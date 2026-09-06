package org.apache.poi.sl.draw;

import io.github.atlan77c.pptx2picture.internal.ForcedLineBreakFixer;
import org.apache.poi.xslf.usermodel.XSLFGroupShape;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTable;
import org.apache.poi.xslf.usermodel.XSLFTableCell;
import org.apache.poi.xslf.usermodel.XSLFTableRow;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xslf.usermodel.XSLFTextShape;

import java.awt.Graphics2D;
import java.awt.geom.Rectangle2D;
import java.util.ArrayList;
import java.util.List;

/**
 * Detecte, pour une forme de texte deja passee par le pipeline reel complet
 * (y compris le retrecissement vertical existant, voir point 2 ci-dessous),
 * la (les) ligne(s) en depassement horizontal reel, puis appelle {@link
 * ForcedLineBreakFixer} pour inserer un saut de ligne force juste avant les
 * derniers mots en trop - en augmentant progressivement le nombre de mots
 * retires (1, puis 2, ...) jusqu'a ce que la ligne corrigee tienne (ou
 * abandon apres quelques essais). Voir Javadoc de {@link ForcedLineBreakFixer}
 * pour la technique d'insertion elle-meme, et {@link
 * io.github.atlan77c.pptx2picture.internal.OverflowAwareTextFitter} pour le
 * correctif complementaire (retrecissement de police) applique en amont dans
 * le pipeline.
 *
 * <p><b>Pourquoi ce paquetage {@code org.apache.poi.sl.draw} plutot que {@code
 * io.github.atlan77c.pptx2picture.internal}</b> (paquetage de tous les autres
 * correctifs de ce projet) : la mesure precise ligne par ligne necessaire ici
 * ({@code DrawFactory.getInstance(graphics).getDrawable(paragraph)}, qui
 * retourne un {@code DrawTextParagraph}, puis son {@code breakText(graphics)}
 * et sa liste {@code lines} de {@code DrawTextFragment}) requiert un acces a
 * la classe {@code DrawTextParagraph} d'Apache POI, qui est package-private -
 * impossible a importer depuis un autre paquetage. Placer cette classe dans
 * le MEME paquetage qu'elle (technique de "paquetage partage", standard en
 * Java sur un classpath classique - non applicable avec des modules JPMS,
 * mais ce projet n'en utilise pas) est la seule facon d'y acceder sans
 * reflexion. Validee par execution reelle en bac a sable le 2026-09-06 avant
 * ce portage (voir {@code bug_debordement_horizontal_placeholder_sans_taille_locale.md}).
 *
 * <p><b>Deux regles d'exclusion</b>, confirmees par l'utilisateur comme des
 * faux positifs de la detection de depassement (voir {@link #isExcluded}) :
 * une petite forme dont les marges internes (insets) dominent sa largeur
 * (ex. un connecteur), et un badge specifique ("Vision cible de la solution
 * EPIC 4") dont le mecanisme de faux positif n'a pas ete elucide (texte gras,
 * hypothese non verifiee de sous-estimation Java2D - voir {@code
 * OverflowAwareTextFitter#maxBoldWidthPremium}).
 */
public final class HorizontalOverflowLineBreakFixer {

    private HorizontalOverflowLineBreakFixer() {
    }

    private static final int MAX_WORDS_TO_DROP = 8;
    private static final int MAX_LINES_FIXED_PER_SHAPE = 10;

    /**
     * Meme marge de securite que {@code WIDTH_SAFETY_MARGIN} deja utilisee
     * ailleurs dans ce projet ({@link
     * io.github.atlan77c.pptx2picture.internal.OverflowAwareTextFitter}) pour
     * le meme ecart documente entre la largeur mesuree par Java2D/breakText()
     * et la largeur reellement dessinee - decouverte et confirmee
     * empiriquement ici le 2026-09-06 (slide 57 : {@code dp.lines}/{@code
     * getWrappingWidth()} indiquait la ligne "tenait" avec 14pt de marge apres
     * avoir retire seulement "et", mais le rendu reel continuait a couper le
     * dernier caractere du mot precedent - "perçus" affiche "perçu" - jusqu'a
     * retirer aussi ce mot). Sans cette marge, la verification par
     * re-breakText() SEULE ne suffit PAS a detecter ce cas (meme outil de
     * mesure biaise des deux cotes).
     */
    private static final double WIDTH_SAFETY_MARGIN = 0.97;

    /** Voir Javadoc de {@link #isExcluded} - regle d'exclusion "petite forme, marges internes dominantes". */
    private static final double INSET_DOMINANCE_THRESHOLD = 0.5;

    /** Voir Javadoc de {@link #isExcluded} - regle d'exclusion "badge Vision cible". */
    private static final String VISION_CIBLE_BADGE_TEXT = "Vision cible de la solution EPIC 4";

    /**
     * Corrige, in-place, le debordement horizontal de toutes les formes de
     * texte du slide (y compris a l'interieur des groupes et des cellules de
     * tableau - voir {@link #collectTextShapes}), par insertion de saut(s) de
     * ligne force(s). A appeler apres tous les autres correctifs de mise en
     * page du pipeline (en particulier {@code
     * OverflowAwareTextFitter#fitOverflowingText}, dont le retrecissement
     * vertical peut reveler un nouveau depassement horizontal par reflow -
     * voir Javadoc de {@link ForcedLineBreakFixer}), et avant {@code
     * slide.draw(graphics)}.
     *
     * @return le nombre de formes de texte effectivement corrigees (au moins
     *         un saut de ligne force insere).
     */
    public static int fixHorizontalOverflow(XSLFSlide slide, Graphics2D graphics) {
        int shapesFixed = 0;
        for (XSLFTextShape ts : collectTextShapes(slide.getShapes())) {
            if (isExcluded(ts)) {
                continue;
            }
            int inserted = fixShape(ts, graphics);
            if (inserted > 0) {
                shapesFixed++;
            }
        }
        return shapesFixed;
    }

    /**
     * Deux faux positifs CONFIRMES PAR L'UTILISATEUR le 2026-09-06 (voir
     * {@code bug_debordement_horizontal_placeholder_sans_taille_locale.md}) :
     * <ol>
     *   <li>une petite forme (ex. un connecteur) dont les marges internes
     *       (insets gauche+droite) consomment a elles seules au moins la
     *       moitie de la largeur de l'ancre - {@code getWrappingWidth()} ne
     *       reflete alors pas le comportement reel de PowerPoint/POI pour
     *       d'aussi petites formes ;</li>
     *   <li>le badge "Vision cible de la solution EPIC 4" (texte gras),
     *       exclu ponctuellement par correspondance exacte de contenu -
     *       mecanisme du faux positif non elucide, jamais generalise en
     *       regle plus large.</li>
     * </ol>
     */
    private static boolean isExcluded(XSLFTextShape ts) {
        Rectangle2D anchor = ts.getAnchor();
        double insetRatio = anchor != null && anchor.getWidth() > 0
                ? (ts.getLeftInset() + ts.getRightInset()) / anchor.getWidth() : 0;
        if (insetRatio >= INSET_DOMINANCE_THRESHOLD) {
            return true;
        }
        return isVisionCibleBadge(ts);
    }

    private static boolean isVisionCibleBadge(XSLFTextShape ts) {
        StringBuilder full = new StringBuilder();
        for (XSLFTextParagraph p : ts.getTextParagraphs()) {
            for (XSLFTextRun r : p.getTextRuns()) {
                String t = r.getRawText();
                if (t != null) {
                    full.append(t);
                }
            }
        }
        return VISION_CIBLE_BADGE_TEXT.equals(full.toString().trim());
    }

    /** @return le nombre de sauts de ligne forces inseres sur cette forme. */
    private static int fixShape(XSLFTextShape ts, Graphics2D graphics) {
        int totalInserted = 0;
        for (int guard = 0; guard < MAX_LINES_FIXED_PER_SHAPE; guard++) {
            Overflow found = findFirstOverflow(ts, graphics);
            if (found == null) {
                break;
            }
            int bestWordsToDrop = -1;
            for (int wordsToDrop = 1; wordsToDrop <= MAX_WORDS_TO_DROP; wordsToDrop++) {
                ForcedLineBreakFixer.SplitPoint sp =
                        ForcedLineBreakFixer.locateSplitPoint(found.paragraph, found.lineText, wordsToDrop);
                if (sp == null) {
                    break; // plus assez de mots dans la ligne : abandon.
                }
                double previewWidth = ForcedLineBreakFixer.previewTruncatedWidth(found.paragraph, sp,
                        () -> remeasureLastLineWidth(found.paragraph, graphics, found.isFirstParagraphOfShape));
                if (previewWidth <= found.wrapWidth * WIDTH_SAFETY_MARGIN) {
                    bestWordsToDrop = wordsToDrop;
                    break;
                }
            }
            if (bestWordsToDrop < 0) {
                break; // aucun nombre de mots retires (jusqu'au maximum) ne suffit : abandon.
            }
            ForcedLineBreakFixer.SplitPoint sp =
                    ForcedLineBreakFixer.locateSplitPoint(found.paragraph, found.lineText, bestWordsToDrop);
            if (sp == null || !ForcedLineBreakFixer.applySplit(found.paragraph, sp)) {
                break;
            }
            totalInserted++;
        }
        return totalInserted;
    }

    /**
     * Relit, apres une troncature temporaire du run de coupure ET un
     * "videment" temporaire de tout ce qui suit (voir Javadoc de {@link
     * ForcedLineBreakFixer#previewTruncatedWidth}), la largeur de la
     * DERNIERE ligne non vide du paragraphe d'essai - sans ambiguite possible
     * puisque rien ne suit ce point dans ce texte d'essai, contrairement a
     * "toujours la ligne 0" (bug corrige le 2026-09-06, voir Javadoc de
     * previewTruncatedWidth).
     */
    private static double remeasureLastLineWidth(XSLFTextParagraph para, Graphics2D graphics, boolean isFirstParagraphOfShape) {
        DrawFactory fact = DrawFactory.getInstance(graphics);
        DrawTextParagraph dp = fact.getDrawable(para);
        dp.setFirstParagraph(isFirstParagraphOfShape);
        dp.breakText(graphics);
        for (int i = dp.lines.size() - 1; i >= 0; i--) {
            String s = dp.lines.get(i).getString();
            if (s != null && !s.trim().isEmpty()) {
                return dp.lines.get(i).getWidth();
            }
        }
        return 0;
    }

    private static Overflow findFirstOverflow(XSLFTextShape ts, Graphics2D graphics) {
        DrawFactory fact = DrawFactory.getInstance(graphics);
        boolean firstParagraphOfShape = true;
        for (XSLFTextParagraph p : ts.getTextParagraphs()) {
            if (rawText(p).trim().isEmpty()) {
                continue;
            }
            boolean isFirst = firstParagraphOfShape;
            DrawTextParagraph dp = fact.getDrawable(p);
            dp.setFirstParagraph(isFirst);
            firstParagraphOfShape = false;
            dp.breakText(graphics);
            for (int i = 0; i < dp.lines.size(); i++) {
                DrawTextFragment line = dp.lines.get(i);
                if (line.getString() == null || line.getString().trim().isEmpty()) {
                    continue;
                }
                double wrapWidth = dp.getWrappingWidth(i == 0, graphics);
                // Marge de securite (voir Javadoc WIDTH_SAFETY_MARGIN) appliquee des la
                // DETECTION initiale, pas seulement lors de la reverification apres coupure -
                // une ligne mesuree "de justesse" en dessous de wrapWidth peut deja etre
                // reellement coupee au rendu (meme phenomene, meme outil de mesure biaise).
                double excess = line.getWidth() - wrapWidth * WIDTH_SAFETY_MARGIN;
                if (excess > 0) {
                    Overflow o = new Overflow();
                    o.paragraph = p;
                    o.lineText = line.getString();
                    o.excess = excess;
                    o.wrapWidth = wrapWidth;
                    o.isFirstParagraphOfShape = isFirst;
                    return o;
                }
            }
        }
        return null;
    }

    private static String rawText(XSLFTextParagraph p) {
        StringBuilder sb = new StringBuilder();
        for (XSLFTextRun r : p.getTextRuns()) {
            String t = r.getRawText();
            if (t != null) {
                sb.append(t);
            }
        }
        return sb.toString();
    }

    /**
     * Parcourt les formes du slide, y compris a l'interieur des groupes ET
     * des cellules de tableau, et ne garde que celles porteuses de texte -
     * meme methode que {@link
     * io.github.atlan77c.pptx2picture.internal.SymbolFontRunFixer}.
     */
    private static List<XSLFTextShape> collectTextShapes(List<XSLFShape> shapes) {
        List<XSLFTextShape> result = new ArrayList<>();
        for (XSLFShape shape : shapes) {
            if (shape instanceof XSLFGroupShape) {
                result.addAll(collectTextShapes(((XSLFGroupShape) shape).getShapes()));
            } else if (shape instanceof XSLFTable) {
                for (XSLFTableRow row : ((XSLFTable) shape).getRows()) {
                    for (XSLFTableCell cell : row.getCells()) {
                        result.add(cell);
                    }
                }
            } else if (shape instanceof XSLFTextShape) {
                result.add((XSLFTextShape) shape);
            }
        }
        return result;
    }

    private static class Overflow {
        XSLFTextParagraph paragraph;
        String lineText;
        double excess;
        double wrapWidth;
        boolean isFirstParagraphOfShape;
    }
}
