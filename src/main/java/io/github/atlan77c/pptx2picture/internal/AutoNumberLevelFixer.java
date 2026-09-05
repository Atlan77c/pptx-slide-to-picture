package io.github.atlan77c.pptx2picture.internal;

import org.apache.poi.sl.draw.DrawFactory;
import org.apache.poi.sl.draw.DrawTextParagraph;
import org.apache.poi.sl.draw.DrawTextShape;
import org.apache.poi.sl.draw.Drawable;
import org.apache.poi.sl.draw.geom.CustomGeometry;
import org.apache.poi.sl.draw.geom.Outline;
import org.apache.poi.sl.usermodel.TextParagraph;
import org.apache.poi.sl.usermodel.TextParagraph.BulletStyle;
import org.apache.poi.sl.usermodel.TextRun;
import org.apache.poi.sl.usermodel.TextShape;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.awt.Graphics2D;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.Collection;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

/**
 * Corrige la numerotation automatique ({@code <a:buAutoNum>}, ex. "1.", "2."...)
 * d'une liste a plusieurs niveaux d'indentation (ex. le sommaire d'un document
 * de cadrage : items de niveau 0 numerotes 1/2/3, chacun suivi d'items de
 * niveau 1 numerotes dans une sequence INDEPENDANTE qui redemarre a 1 sous
 * chaque parent).
 *
 * <p><b>Root cause, confirmee par lecture du code source reel d'Apache POI
 * 5.2.5</b> (module {@code poi}, miroir GitHub officiel, tag {@code
 * REL_5_2_5}) : {@code org.apache.poi.sl.draw.DrawTextShape#drawParagraphs
 * (Graphics2D, double, double)} calcule l'index de numerotation avec un
 * compteur UNIQUE et PLAT, partage par TOUS les paragraphes de la forme,
 * sans jamais lire {@code TextParagraph#getIndentLevel()} :
 * <pre>{@code
 * for (int autoNbrIdx=0; paragraphs.hasNext(); autoNbrIdx++){
 *     TextParagraph p = paragraphs.next();
 *     BulletStyle bs = p.getBulletStyle();
 *     if (bs == null || bs.getAutoNumberingScheme() == null) {
 *         autoNbrIdx = -1;
 *     } else {
 *         Integer startAt = bs.getAutoNumberingStartAt();
 *         if (startAt == null) startAt = 1;
 *         // TODO: handle reset auto number indexes
 *         if (startAt > autoNbrIdx) autoNbrIdx = startAt;
 *     }
 *     dp.setAutoNumberingIdx(autoNbrIdx);
 *     ...
 * }
 * }</pre>
 * Le commentaire {@code // TODO: handle reset auto number indexes} est
 * present tel quel dans le code source POI 5.2.5 : confirmation directe qu'il
 * s'agit d'une limitation connue et non traitee par POI, pas d'un cas non
 * couvert par ce projet. Consequence exacte, verifiee sur un fichier reel
 * (sommaire a 2 niveaux) : chaque item de niveau 0 se retrouve numerote "1."
 * (par coincidence, quand il est precede d'un paragraphe sans puce qui remet
 * le compteur plat a zero), et les items de niveau 1 en dessous poursuivent
 * la meme sequence plate au lieu de redemarrer (2., 3., 4... au lieu de 1.,
 * 2., 3...). {@code buAutoNum/@startAt} ("A partir de" dans le volet Puces et
 * numeros de PowerPoint) n'est utilise QUE comme plancher - une fois le
 * compteur plat deja au-dessus de cette valeur, l'override est silencieusement
 * ignore, et surtout il "contamine" tout ce qui suit (les enfants d'un
 * paragraphe dont le {@code startAt} a saute a 4 continuent eux aussi a
 * partir de 4, au lieu de redemarrer independamment a 1 - confirme sur un
 * second fichier reel construit par l'utilisateur pour tester precisement
 * ce cas).
 *
 * <p><b>{@code DrawTextParagraph#getBullet(...)}</b> (formatage du texte
 * affiche, "1.", "2."...) n'a aucun role dans le bug : il se contente
 * d'appeler {@code AutoNumberingScheme#format(int)} avec l'index qu'on lui
 * donne - tout se joue dans le calcul de cet index, dans {@code
 * DrawTextShape}.
 *
 * <p><b>Correctif</b> : {@link LevelAwareTextShape}, une sous-classe de
 * {@code DrawTextShape} qui reimplemente {@code drawParagraphs(Graphics2D,
 * double, double)} - meme squelette que la methode POI ci-dessus (positions,
 * espacements avant/apres paragraphe, gestion de la premiere ligne : copies a
 * l'identique), seul le calcul de l'index change, delegue a {@link
 * LevelNumbering} qui maintient un compteur INDEPENDANT par niveau
 * d'indentation plutot qu'un compteur plat.
 *
 * <p><b>Semantique retenue pour {@code startAt} - CORRIGEE et CONFIRMEE le
 * 2026-09-05 sur un rendu reel</b> : {@code startAt} est un PLANCHER, jamais
 * un ecrasement inconditionnel - exactement comme le compteur plat de POI
 * (voir l'extrait de code source ci-dessus, {@code if (startAt > autoNbrIdx)
 * autoNbrIdx = startAt;}), mais applique au compteur du SEUL niveau concerne
 * plutot qu'a un compteur global. Une premiere version de ce correctif
 * ecrasait inconditionnellement la valeur par {@code startAt} des qu'il etait
 * present ; l'utilisateur a signale, sur le fichier reel de diagnostic (slide
 * "Sommaire", niveau 0 - "Section A"=1, "Section B" et "Section C" tous
 * deux avec {@code startAt="4"} explicite dans le XML) que le rendu attendu
 * de PowerPoint est 1, 4, 5 (pas 1, 4, 4) : le deuxieme {@code startAt="4"}
 * doit ceder devant la continuation naturelle du niveau 0 (qui vaudrait 5 a
 * ce point, puisque le niveau 0 est deja a 4), alors qu'un ecrasement
 * inconditionnel produisait bien "4" pour les deux paragraphes. Desormais, la
 * valeur naturelle (continuation/reprise/premiere visite - memes regles que
 * sans {@code startAt}, voir {@link LevelNumbering}) est TOUJOURS calculee, et
 * la valeur finale est {@code Math.max(startAt, valeurNaturelle)} - un {@code
 * startAt} n'a d'effet que lorsqu'il depasse ce que la sequence aurait de
 * toute facon produit. Angle mort restant, toujours non verifie faute de
 * PowerPoint/LibreOffice cote agent : un {@code startAt} strictement INFERIEUR
 * a la valeur naturelle (ex. un "redemarrer a 1" explicite sur un niveau deja
 * a 4) reste ignore par ce plancher - aucun fichier reel examine a ce jour
 * n'expose ce cas precis.
 *
 * <p><b>Pourquoi de la reflexion ({@code breakText}/{@code
 * setFirstParagraph})</b> : {@code DrawTextParagraph#breakText(Graphics2D)}
 * et {@code DrawTextParagraph#setFirstParagraph(boolean)} sont {@code
 * protected} et {@code DrawTextParagraph} est une classe FINALE (non
 * sous-classable) - inaccessibles depuis un paquetage tiers autrement que par
 * reflexion (l'astuce "meme paquetage que POI", utilisee pour d'autres
 * membres {@code protected} ailleurs dans ce projet via heritage direct de
 * {@code DrawTextShape}, ne s'applique pas ici car ce sont des membres d'un
 * objet DIFFERENT - {@code DrawTextParagraph}, pas {@code DrawTextShape} -
 * l'acces {@code protected} de Java ne s'etend jamais a un objet tiers non
 * apparente par heritage). Alternative ecartee - dupliquer ce paquetage dans
 * l'arborescence de ce projet (comme {@code AutoShapeGeometryFixer} le fait
 * par heritage direct) : viable uniquement pour une APPLICATION, dangereux
 * pour cette LIBRAIRIE reutilisable (deux jars declarant des types dans le
 * meme paquetage {@code org.apache.poi.sl.draw} provoque une erreur "split
 * package" chez tout consommateur utilisant le systeme de modules Java/JPMS
 * en aval). Les deux {@link Method} sont resolues une seule fois (bloc
 * statique, {@code setAccessible(true)}), avec echec rapide et explicite
 * (erreur au chargement de la classe) si une future version de POI change la
 * signature de ces deux methodes.
 *
 * <p><b>Composition avec {@link AutoShapeGeometryFixer}</b> : {@link
 * DrawFactoryComposer} route 100% des {@code TextShape} vers ce correctif
 * (jamais vers le {@code DrawTextShape} standard de POI, ni directement vers
 * {@code AutoShapeGeometryFixer}) - la numerotation doit etre corrigee sur
 * TOUTE forme de texte, y compris celles dont la geometrie de remplissage est
 * un simple rectangle (cas le plus frequent, dont {@code AutoShapeGeometryFixer}
 * seul ne s'occupe jamais - voir sa Javadoc). {@link #wrap} recoit en
 * parametre la geometrie deja resolue par {@link
 * AutoShapeGeometryFixer#resolveIfQualifies} (extrait de {@code
 * AutoShapeGeometryFixer} specifiquement pour cette composition, {@code
 * null} si aucune correction de geometrie ne s'applique) et produit l'UNIQUE
 * {@code DrawTextShape} concret retourne, qui applique donc les deux
 * correctifs a la fois (numerotation toujours ; contour de remplissage
 * seulement si {@code resolvedGeometry != null}). Bonus gratuit : {@code
 * DrawTextShape#getTextHeight()} appelle en interne {@code drawParagraphs()}
 * (sur un {@code Graphics2D} factice, pour mesurer) - corriger {@code
 * drawParagraphs()} corrige donc AUSSI la mesure de hauteur utilisee par
 * {@link OverflowAwareTextFitter}, sans point de correction separe a
 * prevoir.
 */
public final class AutoNumberLevelFixer {

    private static final Logger LOG = LoggerFactory.getLogger(AutoNumberLevelFixer.class);

    private static final Method BREAK_TEXT_METHOD;
    private static final Method SET_FIRST_PARAGRAPH_METHOD;

    static {
        try {
            BREAK_TEXT_METHOD = DrawTextParagraph.class.getDeclaredMethod("breakText", Graphics2D.class);
            BREAK_TEXT_METHOD.setAccessible(true);
            SET_FIRST_PARAGRAPH_METHOD = DrawTextParagraph.class.getDeclaredMethod("setFirstParagraph", boolean.class);
            SET_FIRST_PARAGRAPH_METHOD.setAccessible(true);
        } catch (NoSuchMethodException e) {
            // Echec rapide et explicite des le chargement de la classe si une future version de
            // POI renomme/modifie la signature de ces deux methodes internes - voir Javadoc de la
            // classe, section "Pourquoi de la reflexion".
            throw new ExceptionInInitializerError(
                    "Apache POI a change la signature de DrawTextParagraph#breakText(Graphics2D) ou "
                            + "#setFirstParagraph(boolean) - AutoNumberLevelFixer doit etre mis a jour "
                            + "en consequence (voir sa Javadoc, section \"Pourquoi de la reflexion\"). Cause : "
                            + e);
        }
    }

    private AutoNumberLevelFixer() {
    }

    /**
     * A appeler juste avant {@code slide.draw(graphics)} pour utiliser CE
     * correctif seul, independamment de {@link AutoShapeGeometryFixer} - voir
     * {@link DrawFactoryComposer} pour la combinaison des deux, utilisee en
     * production par {@link io.github.atlan77c.pptx2picture.PptxSlideRenderer}.
     *
     * @return la valeur precedente du hint {@code Drawable.DRAW_FACTORY} sur
     * ce {@code Graphics2D} (peut etre {@code null}), a repasser telle quelle
     * a {@link #restoreAfterDraw} une fois {@code slide.draw(graphics)}
     * termine.
     */
    public static Object installBeforeDraw(Graphics2D graphics) {
        Object previous = graphics.getRenderingHint(Drawable.DRAW_FACTORY);
        graphics.setRenderingHint(Drawable.DRAW_FACTORY, createDrawFactory());
        if (LOG.isDebugEnabled()) {
            LOG.debug("DrawFactory de numerotation automatique par niveau installe");
        }
        return previous;
    }

    /**
     * A appeler juste apres {@code slide.draw(graphics)}, avec la valeur
     * renvoyee par {@link #installBeforeDraw} - remet {@code graphics} dans
     * l'etat ou {@link #installBeforeDraw} l'a trouve.
     */
    public static void restoreAfterDraw(Graphics2D graphics, Object previousDrawFactory) {
        graphics.setRenderingHint(Drawable.DRAW_FACTORY, previousDrawFactory);
    }

    /**
     * Cree une nouvelle instance du {@code DrawFactory} de ce correctif, sans
     * l'installer - reserve a l'usage standalone (voir {@link
     * #installBeforeDraw}) ; {@link DrawFactoryComposer} utilise {@link #wrap}
     * directement plutot que ce {@code DrawFactory}, pour composer avec
     * {@link AutoShapeGeometryFixer}.
     */
    static DrawFactory createDrawFactory() {
        return new LevelAwareDrawFactory();
    }

    /**
     * Produit le {@code DrawTextShape} a utiliser pour {@code shape} - toujours
     * avec la numerotation par niveau corrigee, et en plus avec le contour de
     * remplissage {@code resolvedGeometry} si non {@code null} (voir Javadoc
     * de la classe, section "Composition avec AutoShapeGeometryFixer").
     *
     * @param resolvedGeometry la geometrie deja resolue par {@link
     *                          AutoShapeGeometryFixer#resolveIfQualifies}, ou
     *                          {@code null} pour garder le contour rectangulaire
     *                          standard de POI.
     */
    static DrawTextShape wrap(TextShape<?, ?> shape, CustomGeometry resolvedGeometry) {
        return new LevelAwareTextShape(shape, resolvedGeometry);
    }

    private static final class LevelAwareDrawFactory extends DrawFactory {
        @Override
        public DrawTextShape getDrawable(TextShape<?, ?> shape) {
            return wrap(shape, null);
        }
    }

    /**
     * Sous-classe de {@code DrawTextShape} qui applique la numerotation par
     * niveau ({@link #drawParagraphs}) et, le cas echeant, le contour de
     * remplissage resolu par heritage de placeholder ({@link
     * #computeOutlines}, delegue a {@link
     * AutoShapeGeometryFixer#computeResolvedOutlines}).
     */
    static final class LevelAwareTextShape extends DrawTextShape {

        private final CustomGeometry resolvedGeometry;

        LevelAwareTextShape(TextShape<?, ?> shape, CustomGeometry resolvedGeometry) {
            super(shape);
            this.resolvedGeometry = resolvedGeometry;
        }

        @Override
        protected Collection<Outline> computeOutlines(Graphics2D graphics) {
            if (resolvedGeometry == null) {
                return super.computeOutlines(graphics);
            }
            return AutoShapeGeometryFixer.computeResolvedOutlines(graphics, getShape(), resolvedGeometry);
        }

        /**
         * Reimplementation de {@code DrawTextShape#drawParagraphs(Graphics2D,
         * double, double)} (POI 5.2.5) - copie a l'identique pour tout ce qui
         * concerne le positionnement (premiere ligne, espacement avant/apres
         * paragraphe), seul le calcul de l'index de numerotation change (
         * delegue a {@link LevelNumbering} au lieu du compteur plat de POI -
         * voir Javadoc de la classe).
         */
        @Override
        public double drawParagraphs(Graphics2D graphics, double x, double y) {
            DrawFactory fact = DrawFactory.getInstance(graphics);

            double y0 = y;
            Iterator<? extends TextParagraph<?, ?, ? extends TextRun>> paragraphs = getShape().iterator();

            LevelNumbering numbering = new LevelNumbering();
            boolean isFirstLine = true;
            while (paragraphs.hasNext()) {
                TextParagraph<?, ?, ? extends TextRun> p = paragraphs.next();
                DrawTextParagraph dp = fact.getDrawable(p);
                BulletStyle bs = p.getBulletStyle();

                int autoNbrIdx;
                if (bs == null || bs.getAutoNumberingScheme() == null) {
                    // Paragraphe sans puce automatique (ex. espaceur vide entre deux blocs d'un
                    // sommaire) : ne participe a aucun niveau de numerotation, n'affecte donc PAS
                    // l'etat de LevelNumbering (contrairement au compteur plat de POI, que ce cas
                    // remet a zero par effet de bord - voir Javadoc de la classe).
                    autoNbrIdx = -1;
                } else {
                    autoNbrIdx = numbering.next(p.getIndentLevel(), bs.getAutoNumberingStartAt());
                }
                dp.setAutoNumberingIdx(autoNbrIdx);
                invokeBreakText(dp, graphics);

                if (isFirstLine) {
                    y += dp.getFirstLineLeading();
                } else {
                    // le montant d'espace blanc vertical avant le paragraphe
                    Double spaceBefore = p.getSpaceBefore();
                    if (spaceBefore == null) {
                        spaceBefore = 0d;
                    }
                    if (spaceBefore > 0) {
                        // une valeur positive signifie un espacement en pourcentage de la hauteur
                        // de la premiere ligne : plus celle-ci est haute, plus l'espace avant le
                        // paragraphe est grand
                        y += spaceBefore * 0.01 * dp.getFirstLineHeight();
                    } else {
                        // une valeur negative signifie l'espacement absolu en points
                        y += -spaceBefore;
                    }
                }

                dp.setPosition(x, y);
                invokeSetFirstParagraph(dp, isFirstLine);
                isFirstLine = false;

                dp.draw(graphics);
                y += dp.getY();

                if (paragraphs.hasNext()) {
                    Double spaceAfter = p.getSpaceAfter();
                    if (spaceAfter == null) {
                        spaceAfter = 0d;
                    }
                    if (spaceAfter > 0) {
                        // une valeur positive signifie un espacement en pourcentage de la hauteur
                        // de la derniere ligne : plus celle-ci est haute, plus l'espace apres le
                        // paragraphe est grand
                        y += spaceAfter * 0.01 * dp.getLastLineHeight();
                    } else {
                        // une valeur negative signifie l'espacement absolu en points
                        y += -spaceAfter;
                    }
                }
            }
            return y - y0;
        }
    }

    /**
     * Maintient, pour une seule forme de texte parcourue paragraphe par
     * paragraphe dans l'ordre du document, un compteur de numerotation
     * automatique INDEPENDANT par niveau d'indentation - le coeur du
     * correctif (voir Javadoc de la classe pour le detail du bug corrige).
     *
     * <p>Regles (verifiees sur trois fichiers reels - sommaire normal, et deux
     * variantes d'un sommaire de test avec {@code startAt} explicite sur des
     * items de niveau 0) - la valeur NATURELLE (sans tenir compte de {@code
     * startAt}) obeit dans tous les cas a :
     * <ul>
     *   <li>poursuite directe du meme niveau que le paragraphe numerote
     *       precedent : incremente ce niveau depuis sa derniere valeur ;</li>
     *   <li>on entre plus profond que le paragraphe numerote precedent :
     *       redemarre ce niveau a 1, meme s'il avait deja une valeur laissee
     *       par une excursion precedente desormais abandonnee ;</li>
     *   <li>on revient a un niveau moins profond (ou egal) apres avoir ete
     *       plus profond : REPREND la sequence de ce niveau la ou elle en
     *       etait (le niveau plus profond n'interrompt que temporairement le
     *       niveau moins profond, il ne le reinitialise pas).</li>
     * </ul>
     * Un {@code startAt} explicite ne fait JAMAIS que relever cette valeur
     * naturelle (plancher, {@code Math.max(startAt, naturelle)}) - voir
     * Javadoc de la classe, section "Semantique retenue pour startAt", pour
     * l'exemple reel qui a impose ce choix. Dans tous les cas, une fois la
     * valeur finale posee, les niveaux plus profonds que celui qu'on vient de
     * fixer sont oublies (ils redemarreront a leur prochaine visite).
     *
     * <p>Un paragraphe SANS puce automatique (voir l'appelant, {@link
     * LevelAwareTextShape#drawParagraphs}) n'appelle jamais {@link #next} -
     * il n'affecte donc aucun etat ici, contrairement au compteur plat de POI
     * qu'un tel paragraphe remet a zero par effet de bord.
     */
    static final class LevelNumbering {

        private final Map<Integer, Integer> counters = new HashMap<>();
        private int previousLevel = Integer.MIN_VALUE;

        /**
         * @param level   niveau d'indentation ({@code TextParagraph#getIndentLevel()})
         *                du paragraphe numerote courant.
         * @param startAt {@code BulletStyle#getAutoNumberingStartAt()} de ce
         *                paragraphe, ou {@code null} si absent.
         * @return l'index a passer a {@code DrawTextParagraph#setAutoNumberingIdx(int)}
         * pour ce paragraphe.
         */
        int next(int level, Integer startAt) {
            boolean continuation = (previousLevel == level);
            int natural;
            if (continuation) {
                natural = counters.getOrDefault(level, 0) + 1;
            } else if (previousLevel > level && counters.containsKey(level)) {
                // on remonte a un niveau deja visite : on reprend sa sequence, on ne la
                // reinitialise pas.
                natural = counters.get(level) + 1;
            } else {
                // premiere visite de ce niveau, ou on vient d'y descendre pour la 1ere fois
                // depuis le dernier paragraphe numerote (meme si ce niveau avait deja une
                // valeur perimee provenant d'une excursion abandonnee).
                natural = 1;
            }
            // startAt n'est qu'un PLANCHER (voir Javadoc de la classe, section "Semantique
            // retenue pour startAt") : il ne fait jamais regresser la sequence naturelle du
            // niveau, il ne peut que la relever.
            int value = (startAt != null) ? Math.max(startAt, natural) : natural;
            counters.put(level, value);
            counters.keySet().removeIf(l -> l > level);
            previousLevel = level;
            return value;
        }
    }

    private static void invokeBreakText(DrawTextParagraph dp, Graphics2D graphics) {
        try {
            BREAK_TEXT_METHOD.invoke(dp, graphics);
        } catch (IllegalAccessException | InvocationTargetException e) {
            throw new IllegalStateException("Echec de l'appel reflexif a DrawTextParagraph#breakText(Graphics2D) "
                    + "- voir Javadoc de AutoNumberLevelFixer, section \"Pourquoi de la reflexion\"", e);
        }
    }

    private static void invokeSetFirstParagraph(DrawTextParagraph dp, boolean firstParagraph) {
        try {
            SET_FIRST_PARAGRAPH_METHOD.invoke(dp, firstParagraph);
        } catch (IllegalAccessException | InvocationTargetException e) {
            throw new IllegalStateException("Echec de l'appel reflexif a "
                    + "DrawTextParagraph#setFirstParagraph(boolean) - voir Javadoc de AutoNumberLevelFixer, "
                    + "section \"Pourquoi de la reflexion\"", e);
        }
    }
}
