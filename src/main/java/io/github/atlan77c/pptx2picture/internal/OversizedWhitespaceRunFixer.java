package io.github.atlan77c.pptx2picture.internal;

import org.apache.poi.xslf.usermodel.XSLFGroupShape;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTable;
import org.apache.poi.xslf.usermodel.XSLFTableCell;
import org.apache.poi.xslf.usermodel.XSLFTableRow;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.ArrayList;
import java.util.List;

/**
 * Corrige un ecart de fidelite decouvert sur un fichier reel : un paragraphe
 * dont le texte VISIBLE tient sur une hauteur normale se retrouve mesure -
 * et donc peint - avec une hauteur de ligne exageree, poussant tout le
 * contenu place apres lui plus bas que chez PowerPoint, au point de chevaucher
 * une forme voisine independante.
 *
 * <p><b>Constat sur le fichier reel</b> (slide 10 de
 * "02.CGSI 2023 04 11 - SI Fraude_DIANE-CGSI_11avril2023_v1.0.pptx", signale
 * par l'utilisateur via capture d'ecran) : le bloc de chiffres cles ("817
 * signalements crees", "839 DE et 30 E signales", "247 referents...", "34
 * auditeurs...") chevauchait un post-it rouge place juste en dessous, alors
 * que PowerPoint n'affiche aucun chevauchement. Inspection XML directe du
 * premier paragraphe ("817 signalements crees") : il contient TROIS runs -
 * {@code "817 "} en 20pt gras, {@code "signalements crees"} en 14pt, puis un
 * run supplementaire ne contenant qu'UN SEUL CARACTERE ESPACE en 48pt
 * ({@code <a:rPr sz="4800">}) - vraisemblablement un residu d'une mise en
 * forme anterieure jamais nettoye (texte redimensionne, l'espace de fin de
 * ligne garde son ancienne taille). Confirme empiriquement par l'utilisateur :
 * la suppression de ce seul run corrige entierement le rendu.
 *
 * <p><b>Mecanisme</b> : ce paragraphe utilise un interligne en POURCENTAGE
 * ({@code <a:lnSpc><a:spcPct val="100000"/></a:lnSpc>}, 100% - au contraire
 * des paragraphes suivants du meme bloc, qui utilisent tous un interligne en
 * points ABSOLUS, {@code <a:spcPts>}, insensible a ce probleme). Un interligne
 * en pourcentage est calcule par POI ({@code DrawTextParagraph.draw()}, 100%
 * x hauteur mesuree de la ligne) a partir de la hauteur Java2D/AWT
 * effectivement mesuree pour CETTE ligne - laquelle depend des metriques
 * (ascendant/descendant) de TOUS les caracteres qui la composent, espaces
 * invisibles compris, {@code LineBreakMeasurer}/{@code TextLayout} ne
 * distinguant pas un caractere "encre visible" d'un caractere blanc. Un
 * espace isole a 48pt gonfle donc la hauteur mesuree de cette ligne bien
 * au-dela de ce que son contenu visible (du texte en 20pt/14pt) necessiterait
 * - hypothese que PowerPoint ne reproduit pas a l'identique. Comme c'est le
 * TOUT PREMIER paragraphe du bloc, cet exces se propage a chacun des
 * paragraphes suivants, empiles les uns sous les autres.
 *
 * <p><b>Pourquoi ce cas echappe a {@link OverflowAwareTextFitter}</b> : ce
 * correctif ne se declenche que si le texte mesure d'une forme depasse SA
 * PROPRE ancre. Ici, l'ancre du bloc de chiffres avait ete dessinee bien plus
 * haute que necessaire (pour laisser de la place au post-it rouge pose par-
 * dessus) : le debordement reste donc toujours a l'interieur de cette ancre,
 * sans jamais chevaucher son PROPRE bord - seule une forme voisine
 * independante (le post-it) est touchee. Un cas de figure different de tout
 * ce que les correctifs existants de ce projet couvrent (qui comparent tous
 * une forme a sa propre ancre, jamais a une forme voisine en l'absence de
 * debordement d'ancre).
 *
 * <p><b>Correctif retenu</b> : avant le rendu, pour tout paragraphe dont
 * l'interligne resolu ({@link XSLFTextParagraph#getLineSpacing()}) est en
 * pourcentage (valeur {@code null} ou positive - un interligne absolu,
 * valeur negative, est deja insensible a ce probleme et n'est pas touche), on
 * repere le run visible (texte non entierement blanc) de plus grande taille
 * du paragraphe, puis on ramene a cette taille tout run entierement blanc
 * (espace(s) seul(s), texte non vide mais {@code trim()} vide) dont la
 * taille la depasse. Le texte n'est jamais modifie, seule la taille de
 * police d'un run invisible est corrigee - sans effet visuel sur le rendu,
 * seulement sur la hauteur de ligne mesuree.
 *
 * <p><b>Limite assumee</b> : la comparaison se fait a l'echelle du
 * PARAGRAPHE XML (tous ses runs), pas de la ligne effectivement peinte apres
 * retour a la ligne automatique - pour un paragraphe qui s'etale sur
 * plusieurs lignes visuelles, un run blanc surdimensionne sur une ligne
 * pourrait en theorie etre compare a un run visible situe sur une AUTRE ligne
 * du meme paragraphe. Non rencontre sur les fichiers reels traites a ce jour
 * (le cas confirme tient sur une seule ligne). Un paragraphe entierement
 * depourvu de texte visible (paragraphe-espaceur vide) n'est jamais touche,
 * faute de run visible de reference.
 */
public final class OversizedWhitespaceRunFixer {

    private static final Logger LOG = LoggerFactory.getLogger(OversizedWhitespaceRunFixer.class);

    private OversizedWhitespaceRunFixer() {
    }

    /**
     * Corrige, in-place, les runs blancs surdimensionnes des paragraphes
     * concernes du slide (voir Javadoc de la classe). A appeler avant
     * {@code slide.draw(graphics)}.
     *
     * @return le nombre de runs effectivement corriges.
     */
    public static int fixOversizedWhitespaceRuns(XSLFSlide slide) {
        int fixed = 0;
        for (XSLFShape shape : collectTextShapes(slide.getShapes())) {
            XSLFTextShape ts = (XSLFTextShape) shape;
            for (XSLFTextParagraph para : ts.getTextParagraphs()) {
                fixed += fixParagraph(shape, para);
            }
        }
        return fixed;
    }

    private static int fixParagraph(XSLFShape shape, XSLFTextParagraph para) {
        Double lineSpacing = para.getLineSpacing();
        if (lineSpacing != null && lineSpacing < 0) {
            // Interligne absolu (points fixes) : la hauteur de ligne peinte par POI ne
            // depend pas des metriques Java2D des runs - rien a corriger ici.
            return 0;
        }

        double maxVisibleSize = -1;
        for (XSLFTextRun run : para.getTextRuns()) {
            if (isVisible(run.getRawText())) {
                Double size = run.getFontSize();
                if (size != null && size > maxVisibleSize) {
                    maxVisibleSize = size;
                }
            }
        }
        if (maxVisibleSize < 0) {
            // Aucun texte visible dans ce paragraphe (paragraphe-espaceur vide) : pas de
            // reference a laquelle comparer un eventuel run blanc, on ne touche a rien.
            return 0;
        }

        int fixed = 0;
        for (XSLFTextRun run : para.getTextRuns()) {
            if (!isBlank(run.getRawText())) {
                continue;
            }
            Double size = run.getFontSize();
            if (size != null && size > maxVisibleSize) {
                if (LOG.isDebugEnabled()) {
                    LOG.debug("{} : run blanc a {}pt (> {}pt, le plus grand run visible de ce paragraphe) "
                                    + "ramene a {}pt pour ne pas gonfler la hauteur de ligne mesuree",
                            shape.getShapeName(), size, maxVisibleSize, maxVisibleSize);
                }
                run.setFontSize(maxVisibleSize);
                fixed++;
            }
        }
        return fixed;
    }

    /** Texte non vide et pas entierement compose d'espaces - contribue visuellement au rendu. */
    private static boolean isVisible(String text) {
        return text != null && !text.trim().isEmpty();
    }

    /** Texte non vide mais entierement compose d'espaces - invisible, mais mesure quand meme par Java2D. */
    private static boolean isBlank(String text) {
        return text != null && !text.isEmpty() && text.trim().isEmpty();
    }

    /** Parcourt les formes du slide, y compris a l'interieur des groupes ET des cellules de tableau (voir {@link SymbolFontRunFixer}). */
    private static List<XSLFShape> collectTextShapes(List<XSLFShape> shapes) {
        List<XSLFShape> result = new ArrayList<>();
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
                result.add(shape);
            }
        }
        return result;
    }
}
