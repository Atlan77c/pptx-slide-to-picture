package io.github.atlan77c.pptx2picture.internal;

import org.apache.poi.sl.usermodel.AutoNumberingScheme;
import org.apache.poi.sl.usermodel.TextParagraph.BulletStyle;
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
import java.util.Locale;

/**
 * Corrige un ecart de fidelite decouvert sur un fichier reel : une liste
 * NUMEROTEE automatiquement ({@code <a:buAutoNum type="arabicPeriod"/>},
 * "1.", "2."... ou {@code type="alphaLcPeriod"}, "a.", "b."...) s'affiche
 * avec des PICTOGRAMMES (icones diverses, un different par item) a la place
 * des chiffres/lettres attendus.
 *
 * <p><b>Constat sur le fichier reel</b> (slide 96 d'un document interne
 * reel, une liste a 7 items numerotes "1." a "7." suivis de 2 items
 * "a."/"b.") : AUCUN paragraphe de cette liste ne declare explicitement de
 * police Wingdings/Symbol dans son XML - la plupart n'ont meme aucun
 * {@code <a:buFont>} local (certains ont {@code <a:buFontTx/>}, "utiliser la
 * police du texte") - et pourtant chaque numero se rend comme un pictogramme
 * different.
 *
 * <p><b>Root cause, confirmee par lecture du code source reel d'Apache POI
 * 5.2.5</b> (miroir GitHub officiel, tag {@code REL_5_2_5}) :
 * <ul>
 *   <li>{@code DrawTextParagraph#getBullet(Graphics2D, ...)} (module
 *       {@code poi}, paquet {@code org.apache.poi.sl.draw}) construit
 *       TOUJOURS le texte du numero avec {@code AutoNumberingScheme#format
 *       (autoNbrIdx)} quand un schema auto est actif - jusque-la correct -
 *       mais dessine ensuite ce texte avec la police issue de {@code
 *       BulletStyle#getBulletFont()}, INDEPENDAMMENT du fait que la puce
 *       soit un numero automatique ou un caractere litteral ;</li>
 *   <li>{@code XSLFTextParagraph#getBulletFont()} (module {@code
 *       poi-ooxml}) resout cette police en remontant la chaine d'heritage
 *       pPr du paragraphe -&gt; {@code lstStyle} de la forme -&gt; style de
 *       la forme parente (layout/master), en ne cherchant QUE l'element
 *       {@code <a:buFont>} a chaque niveau ({@code
 *       ParagraphPropertyFetcher}) - sans jamais verifier si un niveau plus
 *       LOCAL a deja tranche le CHOIX de puce (via {@code <a:buAutoNum>},
 *       {@code <a:buFontTx/>} ou {@code <a:buNone/>}). Resultat : un
 *       paragraphe qui ne pose localement AUCUN {@code <a:buFont>} (que ce
 *       soit parce qu'il n'en declare aucun, ou parce qu'il declare {@code
 *       <a:buFontTx/>} - element que ce chercheur ne reconnait meme pas)
 *       continue de remonter jusqu'a trouver un {@code <a:buFont>} defini
 *       ailleurs, MEME SI ce {@code <a:buFont>} etait destine a un tout
 *       autre choix de puce (un {@code <a:buChar>} litteral) a un niveau
 *       ancetre completement independant.</li>
 * </ul>
 * Sur le fichier reel : le layout de la forme (idx=10, "corps") declare, a
 * son {@code lvl1pPr} (niveau d'indentation 0), la puce PAR DEFAUT de
 * PowerPoint pour ce type de placeholder - {@code <a:buFont
 * typeface="Wingdings" .../><a:buChar char="§"/>} (un petit carre noir). Nos
 * paragraphes numerotes, n'ayant aucun {@code <a:buFont>} LOCAL, heritent
 * de ce "Wingdings" au lieu de la police du texte - et Wingdings remappant
 * TOUTE la plage ASCII imprimable en pictogrammes (aucun rapport visuel avec
 * les chiffres/lettres qu'il remplace), chaque numero ("1.", "2.", "a."...)
 * se dessine comme une icone differente. C'est une limitation reelle
 * d'Apache POI 5.2.5 (le choix de puce {@code buNone}/{@code buChar}/{@code
 * buAutoNum} n'est pas traite comme un groupe exclusif qui stoppe
 * l'heritage des proprietes associees, {@code buFont} y compris), distincte
 * de celle deja corrigee par {@link BulletSymbolFontFixer} (qui, elle, ne
 * concerne que les puces a CARACTERE LITTERAL explicitement mal
 * reconnues par le remappage symbole de POI - {@code getBulletCharacter()}
 * n'est meme jamais appele par {@code getBullet()} quand un schema auto est
 * actif, voir extrait de code ci-dessus).
 *
 * <p><b>Correctif retenu</b> : avant le rendu, pour tout paragraphe dont
 * {@code getBulletStyle().getAutoNumberingScheme()} est non {@code null}
 * (numerotation automatique active) ET qui ne declare AUCUN {@code
 * <a:buFont>} local sur son propre {@code <a:pPr>} (donc dont la valeur
 * observee par {@code getBulletFont()} est forcement HERITEE, jamais
 * choisie pour CE numero precis), si cette police heritee correspond
 * (insensible a la casse, apres troncature d'une eventuelle liste de repli)
 * a "Wingdings" ou "Symbol" - les deux seules polices que POI remappe en
 * pictogrammes, donc les deux seules capables de produire ce symptome -, on
 * force localement {@code buFont} a la police du texte du paragraphe
 * ({@code XSLFTextParagraph#getDefaultFontFamily()}, la meme valeur de
 * repli que POI utilise en interne quand {@code getBulletFont()} renvoie
 * {@code null}). Aucun risque de regresser une puce Wingdings deja
 * fonctionnelle : ce correctif ne touche jamais un paragraphe dont le choix
 * de puce resolu est un caractere litteral (domaine de {@link
 * BulletSymbolFontFixer}), ni un paragraphe qui declare explicitement son
 * propre {@code <a:buFont>} (les 2 items "a."/"b." du meme fichier reel,
 * qui posent {@code <a:buFont typeface="+mj-lt"/>} a ce niveau, restent
 * intouches et continuent de se rendre dans la police de theme voulue).
 *
 * <p><b>Limite assumee</b> : comme {@link BulletSymbolFontFixer}, ce
 * correctif ne detecte que les deux polices que POI reconnait pour son
 * remappage symbole (Wingdings/Symbol) - une police symbole heritee a tort
 * mais non reconnue par POI (Webdings, Wingdings 2/3...) se rendrait de
 * toute facon comme du texte Latin normal (pas un pictogramme), un symptome
 * different non couvert ici.
 */
public final class AutoNumberBulletFontFixer {

    private static final Logger LOG = LoggerFactory.getLogger(AutoNumberBulletFontFixer.class);

    private AutoNumberBulletFontFixer() {
    }

    /**
     * Corrige, in-place, le {@code buFont} des paragraphes a numerotation
     * automatique concernes du slide (voir Javadoc de la classe). A appeler
     * avant {@code slide.draw(graphics)} (l'ordre relatif a {@link
     * BulletSymbolFontFixer} n'a pas d'importance : les deux correctifs
     * ciblent des paragraphes mutuellement exclusifs - numerotation
     * automatique pour celui-ci, caractere litteral pour l'autre).
     *
     * @return le nombre de paragraphes effectivement corriges.
     */
    public static int fixAutoNumberBulletFonts(XSLFSlide slide) {
        int fixed = 0;
        for (XSLFShape shape : collectTextShapes(slide.getShapes())) {
            XSLFTextShape ts = (XSLFTextShape) shape;
            for (XSLFTextParagraph para : ts.getTextParagraphs()) {
                String correctedFont = correctedBulletFontIfNeedsFix(para);
                if (correctedFont == null) {
                    continue;
                }
                String before = para.getBulletFont();
                para.setBulletFont(correctedFont);
                fixed++;
                if (LOG.isDebugEnabled()) {
                    LOG.debug("{} : numero automatique herite a tort la police de puce a caractere '{}' "
                                    + "(non declaree localement) - remplacee par la police du texte '{}'",
                            shape.getShapeName(), before, correctedFont);
                }
            }
        }
        return fixed;
    }

    /**
     * Determine si ce paragraphe est une puce a numerotation automatique
     * dont la police heritee est Wingdings/Symbol SANS avoir ete declaree
     * localement (voir Javadoc de la classe). Retourne la police de
     * remplacement a appliquer si c'est le cas, {@code null} sinon.
     */
    private static String correctedBulletFontIfNeedsFix(XSLFTextParagraph para) {
        BulletStyle bulletStyle = para.getBulletStyle();
        if (bulletStyle == null) {
            return null;
        }
        AutoNumberingScheme scheme = bulletStyle.getAutoNumberingScheme();
        if (scheme == null) {
            // Puce a caractere litteral (ou pas de puce) : domaine de BulletSymbolFontFixer, pas
            // de celui-ci - getBullet() n'appelle meme jamais getBulletCharacter() quand un
            // schema auto est actif, voir Javadoc de la classe.
            return null;
        }
        if (para.getXmlObject().isSetPPr() && para.getXmlObject().getPPr().isSetBuFont()) {
            // Police de puce posee explicitement SUR CE PARAGRAPHE (ex. les items "a."/"b." du
            // fichier reel, <a:buFont typeface="+mj-lt"/>) : jamais heritee, donc jamais le
            // symptome corrige ici, quelle que soit la police en question.
            return null;
        }
        String typeface = para.getBulletFont();
        if (typeface == null || typeface.isEmpty()) {
            return null;
        }
        String firstToken = typeface.split(",", 2)[0].trim().toLowerCase(Locale.ROOT);
        if (!"wingdings".equals(firstToken) && !"symbol".equals(firstToken)) {
            return null;
        }
        return para.getDefaultFontFamily();
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
