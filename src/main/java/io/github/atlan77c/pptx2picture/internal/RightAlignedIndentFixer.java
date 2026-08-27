package io.github.atlan77c.pptx2picture.internal;

import org.apache.poi.sl.usermodel.TextParagraph.TextAlign;
import org.apache.poi.xslf.usermodel.XSLFGroupShape;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTable;
import org.apache.poi.xslf.usermodel.XSLFTableCell;
import org.apache.poi.xslf.usermodel.XSLFTableRow;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.ArrayList;
import java.util.List;

/**
 * Corrige un ecart de rendu decouvert sur un fichier reel : un texte aligne a
 * droite (ex. "P.3") qui deborde entierement de sa boite et vient se
 * superposer au debut du texte de la forme voisine, au point de rendre les
 * deux illisibles (ex. "P.3" + "CONTEXTE" affiches comme "P.GONTEXTE").
 *
 * <p><b>Diagnostic</b> : {@code getTextAlign()}, {@code getFontSize()} et
 * {@code getAnchor()} renvoient tous des valeurs parfaitement correctes pour
 * les deux formes - le probleme n'est visible qu'au rendu reel (confirme par
 * instrumentation directe du pipeline de dessin de POI, en interceptant les
 * appels {@code Graphics2D.drawString}). Ce qui a ete observe : le texte
 * aligne a droite ne se TERMINE pas au bord droit de sa boite comme attendu,
 * il commence quasiment a cet endroit - il deborde donc de toute sa largeur
 * vers la droite au lieu de rester dans sa boite.
 *
 * <p><b>Cause racine identifiee</b> (confirmee en isolant les deux formes
 * dans un fichier reduit, puis en testant chaque propriete une a une) : le
 * paragraphe concerne a un {@code marL} (marge gauche) et un {@code indent}
 * (indentation de premiere ligne) non nuls resolus - {@code marL=285750}
 * {@code indent=-285750} (22,5pt) dans le fichier reel, le motif standard
 * PowerPoint pour une liste a puces avec indentation suspendue (la puce dans
 * la zone d'indentation, le texte a {@code marL}). Deux origines observees
 * sur le meme fichier reel pour cette meme valeur : soit heritee jusqu'au
 * style par defaut du masque de diapositive ({@code <p:txStyles><p:bodyStyle>}),
 * quand rien ne la redefinit ni sur le paragraphe, ni sur le {@code lstStyle}
 * de la forme ou du layout ; soit copiee telle quelle localement sur la
 * {@code pPr} du paragraphe (memes valeurs exactes que le masque - tres
 * probablement une "materialisation" automatique par PowerPoint plutot qu'un
 * choix delibere, PowerPoint ne proposant pas de retrait de liste a puces
 * pour du texte centre/aligne a droite dans son interface). Dans les deux
 * cas, Apache POI ne gere correctement ce {@code marL}/{@code indent} que
 * pour l'alignement a GAUCHE (le cas d'usage normal d'une liste a puces, ou
 * {@code indent} negatif compense exactement {@code marL}) - pour un
 * alignement a DROITE (et vraisemblablement CENTRE, non verifie sur un
 * fichier reel a ce jour), ce {@code marL}/{@code indent} perturbe le calcul
 * de largeur utilise pour positionner le texte, produisant le debordement
 * observe. Confirme par un test isolant chaque variable : ni l'alignement
 * explicite, ni le mode d'autofit ({@code spAutoFit}), ni le lien vers un
 * placeholder seuls ne suffisent a expliquer le bug - seul le
 * {@code marL}/{@code indent} non nul, combine a un alignement different de
 * GAUCHE, le reproduit. Fixer {@code marL=0} et {@code indent=0}
 * explicitement sur un tel paragraphe, sans toucher a rien d'autre, suffit a
 * corriger completement le rendu (verifie sur le fichier reel a l'origine du
 * signalement).
 *
 * <p><b>Correctif retenu</b> : pour tout paragraphe dont l'alignement
 * resolu est DROITE ou CENTRE et dont la valeur resolue de {@code marL} ou
 * {@code indent} est non nulle - qu'elle soit heritee ou definie localement,
 * voir ci-dessus pourquoi les deux cas sont traites de la meme facon - on
 * force explicitement {@code marL=0} et {@code indent=0} sur ce paragraphe.
 * Un {@code marL}/{@code indent} de liste a puces n'a de sens que pour du
 * texte aligne a gauche (avec ou sans puce visible) - PowerPoint lui-meme ne
 * permet pas d'en definir un pour du texte centre/aligne a droite depuis son
 * interface standard - le neutraliser dans ce cas n'a donc aucun effet de
 * bord legitime attendu, meme quand la valeur est presente localement dans
 * le fichier.
 *
 * <p><b>Limite assumee</b> : la verification empirique (fichier reel, rendu
 * instrumente) ne couvre que le cas ALIGNEMENT DROITE. Le cas CENTRE est
 * corrige par le meme mecanisme par prudence/coherence (le calcul de largeur
 * perturbe est vraisemblablement partage entre les deux modes d'alignement
 * dans POI), mais n'a pas ete observe sur un fichier reel a ce jour.
 */
public final class RightAlignedIndentFixer {

    private static final Logger LOG = LoggerFactory.getLogger(RightAlignedIndentFixer.class);

    private RightAlignedIndentFixer() {
    }

    /**
     * Corrige, in-place, le {@code marL}/{@code indent} des paragraphes
     * concernes du slide (voir Javadoc de la classe). A appeler avant
     * {@code slide.draw(graphics)}.
     *
     * @return le nombre de paragraphes effectivement corriges.
     */
    public static int fixInheritedIndent(XSLFSlide slide) {
        int fixed = 0;
        for (XSLFShape shape : collectTextShapes(slide.getShapes())) {
            XSLFTextShape ts = (XSLFTextShape) shape;
            for (XSLFTextParagraph para : ts.getTextParagraphs()) {
                if (shouldClearIndent(para)) {
                    double marL = para.getLeftMargin() == null ? 0 : para.getLeftMargin();
                    double indent = para.getIndent() == null ? 0 : para.getIndent();
                    para.setLeftMargin(0.0);
                    para.setIndent(0.0);
                    fixed++;
                    if (LOG.isDebugEnabled()) {
                        LOG.debug("{} : paragraphe aligne {} avec marL={}pt indent={}pt (retrait de liste "
                                        + "a puces sans effet legitime pour cet alignement) -> remis a zero "
                                        + "(voir Javadoc de la classe)",
                                shape.getShapeName(), para.getTextAlign(), marL, indent);
                    }
                }
            }
        }
        return fixed;
    }

    private static boolean shouldClearIndent(XSLFTextParagraph para) {
        TextAlign align = para.getTextAlign();
        if (align != TextAlign.RIGHT && align != TextAlign.CENTER) {
            return false;
        }
        Double marL = para.getLeftMargin();
        Double indent = para.getIndent();
        return (marL != null && marL != 0.0) || (indent != null && indent != 0.0);
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
