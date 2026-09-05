package io.github.atlan77c.pptx2picture.internal;

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
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;

/**
 * Corrige un ecart de fidelite decouvert sur un fichier reel : des puces de
 * liste censees afficher un pictogramme (fleche, coche...) d'une police
 * "symbole" (Wingdings) s'affichent a la place comme la lettre/le signe Latin
 * litteral portant le meme code de caractere (ex. "Ø" au lieu d'une fleche,
 * "q" au lieu d'une coche).
 *
 * <p><b>Constat sur le fichier reel</b> (slide 9 de
 * "02.CGSI 2023 04 11 - SI Fraude_DIANE-CGSI_11avril2023_v1.0.pptx") : deux
 * declarations de puce Wingdings coexistent dans le meme slide, avec un
 * rendu radicalement different -
 * <ul>
 *   <li>{@code <a:buFont typeface="Wingdings"/>} (nom de police exact, seul)
 *       - se rend CORRECTEMENT (pictogramme visible) ;</li>
 *   <li>{@code <a:buFont typeface="Wingdings,Sans-Serif"/>} (liste de police
 *       de repli au format CSS, avec une police de secours accolee) - se
 *       rend comme la lettre Latin litterale du {@code buChar}, DENATUREE.</li>
 * </ul>
 *
 * <p><b>Cause racine</b> (confirmee en lisant le code source d'Apache POI
 * 5.2.5, {@code DrawTextParagraph.getBullet()} -&gt;
 * {@code DrawFontManagerDefault.mapFontCharset()}) : pour dessiner une puce,
 * POI construit son {@code FontInfo} bullet uniquement a partir de la chaine
 * {@code BulletStyle.getBulletFont()} ({@code new DrawFontInfo(buFontStr)}) -
 * l'attribut XML {@code charset} eventuellement present sur {@code <a:buFont>}
 * n'est jamais lu sur ce chemin de code (hypothese initialement envisagee,
 * puis ecartee apres lecture directe du source POI). Le caractere de la puce
 * n'est redirige vers la zone d'usage prive Unicode (decalage +0xF000, la ou
 * vivent les glyphes pictogrammes de Wingdings dans sa table TrueType) que si
 * {@code knownSymbolFonts.contains(typeface)} - un {@code TreeSet} interne a
 * POI, insensible a la casse, ne contenant QUE les deux chaines exactes
 * {@code "Wingdings"} et {@code "Symbol"}. Une valeur comme
 * {@code "Wingdings,Sans-Serif"} ne matche aucune des deux : POI peint alors
 * le caractere comme du texte normal, d'ou la lettre Latin litterale affichee
 * a la place du pictogramme.
 *
 * <p><b>Distinct de {@link SymbolFontRunFixer}</b> : ce dernier corrige les
 * polices symbole posees via {@code <a:sym>} sur des RUNS de texte
 * ({@code <a:rPr>}) - un chemin de rendu POI totalement different de celui
 * des puces de paragraphe ({@code <a:buFont>}/{@code <a:buChar>} dans
 * {@code <a:pPr>}, gere par {@code DrawTextParagraph.getBullet()}). Aucun des
 * deux correctifs ne couvre le cas de l'autre.
 *
 * <p><b>Correctif retenu</b> : avant le rendu, pour tout paragraphe dont la
 * puce resolue (fusion pPr -&gt; lstStyle de la forme -&gt; layout -&gt;
 * master, via {@link XSLFTextParagraph#getBulletFont()}/
 * {@link XSLFTextParagraph#getBulletCharacter()}) porte un {@code buFont}
 * dont le premier element de la liste separee par virgules correspond, sans
 * tenir compte de la casse, exactement a {@code "Wingdings"} ou
 * {@code "Symbol"} (les deux seules polices que POI sait reconnaitre pour ce
 * remappage), on force localement {@code typeface} a ce nom exact seul, sans
 * liste de repli - reproduisant exactement la forme deja presente, et deja
 * fonctionnelle, dans ce meme fichier reel pour les puces qui se rendent
 * correctement.
 *
 * <p><b>Limite assumee</b> : les autres polices symbole Windows courantes
 * (Webdings, Wingdings 2/3, Marlett...) ne sont PAS corrigees par ce fixer -
 * meme parfaitement nommees, POI 5.2.5 ne les reconnait pas du tout pour ce
 * remappage de puce (seules "Wingdings"/"Symbol" sont codees en dur cote
 * POI) ; les corriger demanderait de contourner entierement le pipeline de
 * dessin de puce de POI (point d'extension {@code DrawFactory}, comme
 * {@link ConnectorArrowFixer}/{@link PictureGeometryClipFixer}), pas juste de
 * normaliser le XML source - non implemente, motif non rencontre a ce jour
 * sur les fichiers reels traites.
 */
public final class BulletSymbolFontFixer {

    private static final Logger LOG = LoggerFactory.getLogger(BulletSymbolFontFixer.class);

    /**
     * Polices symbole reconnues par le remappage interne d'Apache POI 5.2.5
     * ({@code DrawFontManagerDefault.knownSymbolFonts}) - cle en minuscules
     * pour une comparaison insensible a la casse, valeur = orthographe
     * canonique a ecrire dans le fichier.
     */
    private static final Map<String, String> CANONICAL_SYMBOL_FONTS = new HashMap<>();

    static {
        CANONICAL_SYMBOL_FONTS.put("wingdings", "Wingdings");
        CANONICAL_SYMBOL_FONTS.put("symbol", "Symbol");
    }

    private BulletSymbolFontFixer() {
    }

    /**
     * Corrige, in-place, le {@code buFont} des paragraphes a puce concernes
     * du slide (voir Javadoc de la classe). A appeler avant
     * {@code slide.draw(graphics)}.
     *
     * @return le nombre de paragraphes effectivement corriges.
     */
    public static int fixSymbolBulletFonts(XSLFSlide slide) {
        int fixed = 0;
        for (XSLFShape shape : collectTextShapes(slide.getShapes())) {
            XSLFTextShape ts = (XSLFTextShape) shape;
            for (XSLFTextParagraph para : ts.getTextParagraphs()) {
                String canonical = canonicalSymbolFontIfNeedsFix(para);
                if (canonical == null) {
                    continue;
                }
                String before = para.getBulletFont();
                para.setBulletFont(canonical);
                fixed++;
                if (LOG.isDebugEnabled()) {
                    LOG.debug("{} : puce '{}' avec buFont='{}' normalise en '{}' (liste de police de repli "
                                    + "non reconnue telle quelle par le remappage de police symbole d'Apache POI)",
                            shape.getShapeName(), para.getBulletCharacter(), before, canonical);
                }
            }
        }
        return fixed;
    }

    /**
     * Determine si le {@code buFont} resolu de ce paragraphe est une des
     * polices symbole reconnues par POI ({@code Wingdings}/{@code Symbol}),
     * declaree avec une liste de repli qui empeche POI de la reconnaitre
     * telle quelle. Retourne l'orthographe canonique a appliquer si c'est le
     * cas, {@code null} sinon (pas de puce caractere, pas de buFont resolu,
     * police non reconnue, ou deja sous sa forme canonique exacte).
     */
    private static String canonicalSymbolFontIfNeedsFix(XSLFTextParagraph para) {
        String bulletChar = para.getBulletCharacter();
        if (bulletChar == null || bulletChar.isEmpty()) {
            return null;
        }
        String typeface = para.getBulletFont();
        if (typeface == null || typeface.isEmpty()) {
            return null;
        }
        String firstToken = typeface.split(",", 2)[0].trim();
        String canonical = CANONICAL_SYMBOL_FONTS.get(firstToken.toLowerCase(Locale.ROOT));
        if (canonical == null || typeface.equals(canonical)) {
            return null;
        }
        return canonical;
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
