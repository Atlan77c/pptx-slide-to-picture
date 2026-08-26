package io.github.atlan77c.pptx2picture.internal;

import org.apache.poi.sl.usermodel.TextShape;
import org.apache.poi.sl.usermodel.VerticalAlignment;
import org.apache.poi.xslf.usermodel.XSLFGroupShape;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.awt.Graphics2D;
import java.awt.geom.Rectangle2D;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Corrige un ecart de fidelite observe entre le rendu natif d'Apache POI
 * ({@code XSLFSlide.draw(Graphics2D)}) et le rendu natif de PowerPoint :
 * pour certaines polices, les metriques verticales calculees par le pipeline
 * Java2D/AWT surestiment la hauteur de texte reellement necessaire (jusqu'a
 * 30-35% observe dans certains cas) par rapport a ce que produit le moteur de
 * rendu de PowerPoint. POI n'applique par ailleurs jamais le retrecissement
 * automatique de police ("shrink text on overflow", stocke par PowerPoint
 * dans {@code <a:normAutofit fontScale="...">}) lors du rendu :
 * {@code slide.draw()} utilise toujours la taille de police brute.
 *
 * <p>Consequence non corrigee : une zone de texte peut deborder visuellement
 * de sa boite d'origine et chevaucher une autre forme voisine.
 *
 * <p><b>Strategie retenue</b> : on ne "triche" pas sur les metriques de ligne calculees par
 * Java2D (une compensation directe de {@code DrawTextFragment.getHeight()}
 * corrige bien le debordement mais introduit un nouveau chevauchement entre
 * lignes consecutives d'une meme zone de texte, car cette valeur sert aussi
 * de pas d'avancement vertical entre lignes dans POI). On reduit a la place
 * la taille de police reelle des formes concernees, ce qui reduit glyphes et
 * interlignage ensemble et proportionnellement - par construction, cela ne
 * peut jamais faire chevaucher deux lignes entre elles.
 *
 * <p>Pour eviter de retrecir des formes qui n'en ont pas besoin (la
 * surestimation touche la quasi-totalite des formes d'un fichier utilisant
 * une police concernee, y compris des formes isolees sans aucun risque de
 * collision visuelle reelle), le retrecissement d'une forme marquee
 * {@link TextShape.TextAutofit#NONE} n'est declenche que si son debordement
 * calcule chevaucherait reellement l'emplacement d'une autre forme de texte
 * non vide du meme slide. Les formes {@link TextShape.TextAutofit#NORMAL}
 * (retrecissement automatique voulu par l'auteur du fichier) sont, elles,
 * toujours retrecies en cas de depassement reel - c'est le comportement
 * natif de PowerPoint pour ce type de forme.
 *
 * <p><b>Cas {@link TextShape.TextAutofit#SHAPE}</b> ("redimensionner la forme
 * selon le texte", {@code spAutoFit} en OOXML) : traite comme {@code NORMAL}
 * (retrecissement systematique de la police en cas de depassement mesure).
 * Une premiere version agrandissait plutot la hauteur de l'anchor - fidele au
 * sens strict de ce mode d'autofit dans PowerPoint (qui ne reduit jamais la
 * police, la boite grandit) - mais observee en pratique sur un fichier reel
 * a provoquer de nouveaux chevauchements avec les formes voisines situees en
 * dessous, puisque l'agrandissement deplace la limite basse de la boite sans
 * tenir compte du reste de la mise en page. Retrecir la police laisse toutes
 * les autres formes du slide a leur place d'origine : moins fidele au
 * mecanisme technique de ce mode d'autofit, mais plus fidele au rendu global
 * du diagramme, ce qui est le critere qui compte ici.
 *
 * <p><b>Exception au retrecissement force ({@code NONE})</b> : decouverte sur
 * un fichier reel (une boite {@code noAutofit} nettement plus petite que la
 * taille de police declaree, ne contenant qu'un seul caractere decoratif
 * volontairement surdimensionne par l'auteur pour dessiner un accent visuel -
 * chevauchant deliberement une forme voisine, exactement comme PowerPoint
 * l'affiche). Si la taille de police <em>declaree</em> (avant toute mesure)
 * depasse deja a elle seule la hauteur de la boite, le debordement n'est pas
 * un artefact de mesure Java2D : PowerPoint montrerait le meme debordement
 * quelle que soit la precision du calcul de metriques, puisque meme un calcul
 * parfait ne ferait jamais tenir un texte aussi grand dans une boite aussi
 * petite. Un tel cas n'est donc jamais retreci de force, meme s'il chevauche
 * effectivement une autre forme de texte.
 *
 * <p><b>Deuxieme garde-fou ({@code NONE}) : objectif "sans collision"
 * plutot que "tient dans la boite"</b> - decouverte sur un autre fichier
 * reel (boite {@code noAutofit} contenant plusieurs paragraphes de texte,
 * alignement "Milieu"). Contrairement au cas precedent, aucune taille de
 * police declaree ne depasse a elle seule la hauteur de la boite : c'est le
 * volume de texte (nombre de lignes) qui fait que le retrecissement visant
 * un ajustement complet dans la boite (le meme objectif que pour
 * {@code NORMAL}/{@code SHAPE}) n'atteint jamais sa cible, meme ecrase
 * jusqu'a la limite basse ({@link #MIN_SCALE}) ou {@link #MAX_ITER}
 * iterations. Or PowerPoint n'exige jamais qu'un texte {@code noAutofit}
 * tienne entierement dans sa boite - seulement qu'il ne produise pas de
 * chevauchement genant avec une forme voisine, ce qui est la seule raison
 * d'etre du retrecissement force ici. Quand l'ajustement complet echoue, une
 * deuxieme passe vise donc un objectif plus modeste : ne retrecir que
 * jusqu'a ce que la collision reelle detectee plus haut disparaisse
 * (recalculee a chaque iteration), en acceptant un debordement residuel
 * au-dela de la boite tant qu'il ne chevauche plus la forme voisine -
 * exactement ce que montre PowerPoint pour ce type de boite. Si meme cette
 * cible plus modeste n'est pas atteinte a la limite basse, la taille
 * d'origine est restauree en dernier recours plutot que de produire un
 * texte ecrase qui chevauche quand meme.
 */
public final class OverflowAwareTextFitter {

    private static final Logger LOG = LoggerFactory.getLogger(OverflowAwareTextFitter.class);

    /** Pas de reduction applique a chaque iteration (mimique les paliers utilises par PowerPoint). */
    private static final double STEP = 0.02;
    private static final double MIN_SCALE = 0.25;
    private static final int MAX_ITER = 38;

    /**
     * Marge de securite appliquee a la hauteur cible lors du retrecissement :
     * on arrete de retrecir des que le texte mesure tient dans {@code anchor.getHeight() * SAFETY_MARGIN}
     * plutot que dans {@code anchor.getHeight()} strictement. Sans cette marge, un cas reel
     * (diagnostic sur un fichier de production - forme "spAutoFit" a plusieurs paragraphes, ecart de
     * mesure de ~30% entre Java2D et PowerPoint, dans le haut de la fourchette deja documentee)
     * a montre que le retrecissement s'arretait des que getTextHeight() repassait sous
     * anchor.getHeight(), mais avec une marge residuelle de moins de 1pt sur une hauteur de plusieurs
     * dizaines de points : largement insuffisant pour absorber a la fois l'epaisseur du trait de
     * bordure de la forme et un eventuel ecart residuel entre la mesure de getTextHeight() et ce que
     * slide.draw() peint reellement - le texte continuait donc a chevaucher visuellement la bordure
     * malgre un retrecissement de police tres marque (jusqu'a -30% environ dans ce cas).
     */
    private static final double SAFETY_MARGIN = 0.97;

    private OverflowAwareTextFitter() {
    }

    /**
     * Retrecit, in-place, les formes de texte du slide qui debordent
     * reellement de leur boite (voir Javadoc de la classe). Modifie les
     * tailles de police des {@link XSLFTextRun} concernes ; a appeler avant
     * {@code slide.draw(graphics)}.
     *
     * @return le nombre de formes effectivement retrecies.
     */
    public static int fitOverflowingText(XSLFSlide slide, Graphics2D graphics) {
        int count = 0;
        List<XSLFShape> allTextShapes = collectTextShapes(slide.getShapes());

        for (XSLFShape shape : allTextShapes) {
            XSLFTextShape ts = (XSLFTextShape) shape;
            TextShape.TextAutofit autofit = ts.getTextAutofit();
            // NORMAL et SHAPE partagent la meme logique : PowerPoint ne montre jamais
            // de debordement reel pour ces deux modes, donc tout debordement mesure ici
            // est un artefact du calcul de metriques Java2D (jamais une intention de
            // l'auteur) et est systematiquement corrige. NONE est le seul cas ou un
            // debordement peut etre volontaire (voir "forced" ci-dessous).
            boolean forced = autofit == TextShape.TextAutofit.NONE;

            Rectangle2D anchor = ts.getAnchor();
            if (anchor == null || anchor.getHeight() <= 0) {
                continue;
            }

            if (forced) {
                double maxDeclaredFontSize = maxDeclaredFontSize(ts);
                if (maxDeclaredFontSize > 0 && maxDeclaredFontSize > anchor.getHeight()) {
                    if (LOG.isDebugEnabled()) {
                        LOG.debug("{} : taille de police declaree ({}pt) > hauteur de la boite ({}pt) -> "
                                + "debordement volontaire de l'auteur (ex. gros caractere decoratif), non retrecie",
                                shape.getShapeName(), maxDeclaredFontSize, anchor.getHeight());
                    }
                    continue;
                }

                double initialTextHeight = ts.getTextHeight(graphics);
                VerticalAlignment valign = ts.getVerticalAlignment();
                List<Rectangle2D> overflowZones = computeOverflowZones(anchor, initialTextHeight, valign);
                XSLFShape collidingShape = overflowZones.isEmpty() ? null : findCollidingShape(overflowZones, ts, allTextShapes);
                if (collidingShape == null) {
                    if (!overflowZones.isEmpty() && LOG.isDebugEnabled()) {
                        LOG.debug("{} : debordement mesure ({} > {}) mais aucune collision avec une autre "
                                + "forme de texte -> non retrecie", shape.getShapeName(), initialTextHeight, anchor.getHeight());
                    }
                    continue;
                }
                if (LOG.isDebugEnabled()) {
                    LOG.debug("{} : debordement mesure ({} > {}, alignement {}) entre en collision avec '{}' -> retrecissement force",
                            shape.getShapeName(), initialTextHeight, anchor.getHeight(), valign, collidingShape.getShapeName());
                }
            }

            Map<XSLFTextRun, Double> baseline = captureBaselineFontSizes(ts);
            if (baseline.isEmpty()) {
                continue; // tailles heritees du theme/layout, non modifiables ici
            }

            double factor = 1.0;
            int iter = 0;
            double textHeight = ts.getTextHeight(graphics);
            boolean didShrink = false;
            double targetHeight = anchor.getHeight() * SAFETY_MARGIN;

            while (textHeight > targetHeight && factor > MIN_SCALE && iter < MAX_ITER) {
                factor -= STEP;
                for (Map.Entry<XSLFTextRun, Double> e : baseline.entrySet()) {
                    e.getKey().setFontSize(Math.max(1.0, e.getValue() * factor));
                }
                textHeight = ts.getTextHeight(graphics);
                iter++;
                didShrink = true;
            }

            if (forced && textHeight > targetHeight) {
                // Le retrecissement visant un ajustement complet dans la boite n'a pas
                // converge, meme en ecrasant la police jusqu'a sa limite basse (MIN_SCALE)
                // ou le nombre max d'iterations (MAX_ITER) : la boite est structurellement
                // trop petite pour ce texte (ex. legende/annotation flottante avec bien plus
                // de texte que la boite ne pourra jamais en contenir - meme dans
                // PowerPoint). Deuxieme passe avec un objectif different : PowerPoint
                // n'exige jamais qu'un texte noAutofit tienne entierement dans sa boite -
                // seulement qu'il ne produise pas un chevauchement genant avec une forme
                // voisine, ce qui est la seule raison d'etre du retrecissement force (voir
                // Javadoc de la classe). On repart donc de la taille d'origine et on ne
                // retrecit que jusqu'a la disparition de la collision reelle detectee plus
                // haut - un debordement residuel au-dela de la boite, mais sans chevaucher
                // la forme voisine, est accepte tel quel.
                for (Map.Entry<XSLFTextRun, Double> e : baseline.entrySet()) {
                    e.getKey().setFontSize(e.getValue());
                }
                VerticalAlignment valign = ts.getVerticalAlignment();
                factor = 1.0;
                iter = 0;
                didShrink = false;
                textHeight = ts.getTextHeight(graphics);
                List<Rectangle2D> zones = computeOverflowZones(anchor, textHeight, valign);
                boolean stillColliding = !zones.isEmpty() && findCollidingShape(zones, ts, allTextShapes) != null;

                while (stillColliding && factor > MIN_SCALE && iter < MAX_ITER) {
                    factor -= STEP;
                    for (Map.Entry<XSLFTextRun, Double> e : baseline.entrySet()) {
                        e.getKey().setFontSize(Math.max(1.0, e.getValue() * factor));
                    }
                    textHeight = ts.getTextHeight(graphics);
                    zones = computeOverflowZones(anchor, textHeight, valign);
                    stillColliding = !zones.isEmpty() && findCollidingShape(zones, ts, allTextShapes) != null;
                    iter++;
                    didShrink = true;
                }

                if (stillColliding) {
                    // Meme une reduction jusqu'a la limite basse ne suffit pas a ecarter la
                    // collision : la taille d'origine est restauree plutot que de produire
                    // un texte ecrase qui chevauche quand meme la forme voisine.
                    for (Map.Entry<XSLFTextRun, Double> e : baseline.entrySet()) {
                        e.getKey().setFontSize(e.getValue());
                    }
                    if (LOG.isDebugEnabled()) {
                        LOG.debug("{} : aucun ajustement (complet ou partiel jusqu'a disparition de la collision) "
                                + "n'a ete possible -> boite structurellement trop petite pour ce texte, taille "
                                + "d'origine restauree", shape.getShapeName());
                    }
                    continue;
                }
                if (didShrink && LOG.isDebugEnabled()) {
                    LOG.debug("{} : ajustement complet dans la boite impossible, retrecie a {}% seulement jusqu'a "
                            + "disparition de la collision (hauteur texte {}pt, boite {}pt) - debordement residuel "
                            + "au-dela de la boite accepte, comme le ferait PowerPoint pour ce mode noAutofit",
                            shape.getShapeName(), Math.round(factor * 100), textHeight, anchor.getHeight());
                }
            }

            if (didShrink) {
                count++;
                if (LOG.isDebugEnabled()) {
                    LOG.debug("{} retrecie a {}% (hauteur texte {}pt -> boite {}pt){}",
                            shape.getShapeName(), Math.round(factor * 100), textHeight, anchor.getHeight(),
                            forced ? " [force : autofit=NONE dans le fichier d'origine]" : "");
                }
            }
        }
        return count;
    }

    /**
     * Plus grande taille de police <em>declaree</em> (avant toute reduction)
     * parmi les runs non vides de la forme, ou -1 si aucune n'a de taille
     * explicite (heritee du theme/layout, non determinable ici sans parcourir
     * la chaine d'heritage complete).
     */
    private static double maxDeclaredFontSize(XSLFTextShape ts) {
        double max = -1;
        for (XSLFTextParagraph para : ts.getTextParagraphs()) {
            for (XSLFTextRun run : para.getTextRuns()) {
                String text = run.getRawText();
                if (text == null || text.isEmpty()) {
                    continue;
                }
                Double size = run.getFontSize();
                if (size != null && size > max) {
                    max = size;
                }
            }
        }
        return max;
    }

    private static Map<XSLFTextRun, Double> captureBaselineFontSizes(XSLFTextShape ts) {
        Map<XSLFTextRun, Double> baseline = new HashMap<>();
        for (XSLFTextParagraph para : ts.getTextParagraphs()) {
            for (XSLFTextRun run : para.getTextRuns()) {
                Double size = run.getFontSize();
                if (size != null) {
                    baseline.put(run, size);
                }
            }
        }
        return baseline;
    }

    /**
     * Calcule la ou les zones de "debordement" - au-dela de l'anchor
     * d'origine d'une forme - qu'occuperait son texte s'il n'etait pas
     * retreci, selon son alignement vertical (TOP/MIDDLE/BOTTOM ; TOP par
     * defaut si non precise, convention OOXML {@code anchor="t"}). Retourne
     * une liste vide si le texte ne deborde pas.
     */
    static List<Rectangle2D> computeOverflowZones(Rectangle2D anchor, double textHeight, VerticalAlignment valign) {
        List<Rectangle2D> zones = new ArrayList<>();
        double excess = textHeight - anchor.getHeight();
        if (excess <= 0) {
            return zones;
        }
        VerticalAlignment v = (valign == null) ? VerticalAlignment.TOP : valign;
        switch (v) {
            case BOTTOM:
                zones.add(new Rectangle2D.Double(anchor.getX(), anchor.getY() - excess, anchor.getWidth(), excess));
                break;
            case MIDDLE:
                double half = excess / 2.0;
                zones.add(new Rectangle2D.Double(anchor.getX(), anchor.getY() - half, anchor.getWidth(), half));
                zones.add(new Rectangle2D.Double(anchor.getX(), anchor.getY() + anchor.getHeight(), anchor.getWidth(), half));
                break;
            case TOP:
            default:
                zones.add(new Rectangle2D.Double(anchor.getX(), anchor.getY() + anchor.getHeight(), anchor.getWidth(), excess));
                break;
        }
        return zones;
    }

    /**
     * Indique si l'une des zones de debordement fournies chevauche l'anchor
     * d'une autre forme de texte non vide du slide. Les formes sans texte
     * (rectangles/panneaux de fond) sont ignorees : un debordement dessus
     * n'occasionne aucune confusion visuelle.
     */
    static boolean overflowCollidesWithText(List<Rectangle2D> overflowZones, XSLFTextShape self, List<XSLFShape> allTextShapes) {
        return findCollidingShape(overflowZones, self, allTextShapes) != null;
    }

    /**
     * Comme {@link #overflowCollidesWithText}, mais retourne la forme
     * responsable de la collision (ou {@code null}) plutot qu'un simple
     * booleen - utilise pour enrichir les logs de diagnostic avec le nom de
     * la forme en cause.
     */
    private static XSLFShape findCollidingShape(List<Rectangle2D> overflowZones, XSLFTextShape self, List<XSLFShape> allTextShapes) {
        for (Rectangle2D zone : overflowZones) {
            for (XSLFShape other : allTextShapes) {
                if (other == self || !(other instanceof XSLFTextShape)) {
                    continue;
                }
                XSLFTextShape ots = (XSLFTextShape) other;
                String text = ots.getText();
                if (text == null || text.trim().isEmpty()) {
                    continue;
                }
                Rectangle2D otherAnchor = ots.getAnchor();
                if (otherAnchor != null && zone.intersects(otherAnchor)) {
                    return other;
                }
            }
        }
        return null;
    }

    /** Parcourt les formes du slide, y compris a l'interieur des groupes, et ne garde que celles porteuses de texte. */
    private static List<XSLFShape> collectTextShapes(List<XSLFShape> shapes) {
        List<XSLFShape> result = new ArrayList<>();
        for (XSLFShape shape : shapes) {
            if (shape instanceof XSLFGroupShape) {
                result.addAll(collectTextShapes(((XSLFGroupShape) shape).getShapes()));
            } else if (shape instanceof XSLFTextShape) {
                result.add(shape);
            }
        }
        return result;
    }
}
