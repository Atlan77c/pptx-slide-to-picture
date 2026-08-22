package io.github.atlan77c.pptx2image.internal;

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
 * pour certaines polices (ex. PoliceX, la police de l'Etat francais), les
 * metriques verticales calculees par le pipeline Java2D/AWT surestiment la
 * hauteur de texte reellement necessaire d'environ 30 a 35% par rapport a ce
 * que produit le moteur de rendu de PowerPoint. POI n'applique par ailleurs
 * jamais le retrecissement automatique de police ("shrink text on overflow",
 * stocke par PowerPoint dans {@code <a:normAutofit fontScale="...">}) lors du
 * rendu : {@code slide.draw()} utilise toujours la taille de police brute.
 *
 * <p>Consequence non corrigee : une zone de texte peut deborder visuellement
 * de sa boite d'origine et chevaucher une autre forme voisine.
 *
 * <p><b>Strategie retenue</b> (voir {@code conversion_pptx_vers_images.md} du
 * projet d'origine pour l'historique complet des approches essayees et
 * ecartees) : on ne "triche" pas sur les metriques de ligne calculees par
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
 */
public final class OverflowAwareTextFitter {

    private static final Logger LOG = LoggerFactory.getLogger(OverflowAwareTextFitter.class);

    /** Pas de reduction applique a chaque iteration (mimique les paliers utilises par PowerPoint). */
    private static final double STEP = 0.02;
    private static final double MIN_SCALE = 0.25;
    private static final int MAX_ITER = 38;

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
            boolean forced = autofit == TextShape.TextAutofit.NONE;
            if (autofit != TextShape.TextAutofit.NORMAL && !forced) {
                // TextAutofit.SHAPE (la boite grandit pour epouser le texte) : rien
                // a retrecir, le texte ne "deborde" jamais par construction. POI ne
                // fait pas non plus grandir reellement la boite au rendu, mais ce
                // n'est pas le probleme traite ici.
                continue;
            }

            Rectangle2D anchor = ts.getAnchor();
            if (anchor == null || anchor.getHeight() <= 0) {
                continue;
            }

            if (forced) {
                double initialTextHeight = ts.getTextHeight(graphics);
                List<Rectangle2D> overflowZones = computeOverflowZones(anchor, initialTextHeight, ts.getVerticalAlignment());
                boolean collides = !overflowZones.isEmpty() && overflowCollidesWithText(overflowZones, ts, allTextShapes);
                if (!collides) {
                    if (!overflowZones.isEmpty() && LOG.isDebugEnabled()) {
                        LOG.debug("{} : debordement mesure ({} > {}) mais aucune collision avec une autre "
                                + "forme de texte -> non retrecie", shape.getShapeName(), initialTextHeight, anchor.getHeight());
                    }
                    continue;
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

            while (textHeight > anchor.getHeight() && factor > MIN_SCALE && iter < MAX_ITER) {
                factor -= STEP;
                for (Map.Entry<XSLFTextRun, Double> e : baseline.entrySet()) {
                    e.getKey().setFontSize(Math.max(1.0, e.getValue() * factor));
                }
                textHeight = ts.getTextHeight(graphics);
                iter++;
                didShrink = true;
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
                    return true;
                }
            }
        }
        return false;
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
