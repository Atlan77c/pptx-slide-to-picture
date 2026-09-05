package io.github.atlan77c.pptx2picture.internal;

import org.apache.poi.common.usermodel.fonts.FontGroup;
import org.apache.poi.sl.usermodel.PaintStyle;
import org.apache.poi.xslf.usermodel.XSLFGroupShape;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTable;
import org.apache.poi.xslf.usermodel.XSLFTableCell;
import org.apache.poi.xslf.usermodel.XSLFTableRow;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.openxmlformats.schemas.drawingml.x2006.main.CTRegularTextRun;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextCharacterProperties;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextFont;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.ArrayList;
import java.util.List;

/**
 * Corrige un ecart de fidelite decouvert lors d'un test de masse sur un
 * document reel : certains runs de texte portent un {@code <a:sym typeface="...">}
 * (utilise par PowerPoint pour un caractere isole insere via Insert &gt; Symbole,
 * typiquement une fleche Wingdings, ex. U+F0E8) alors que leur texte contient
 * AUSSI du texte normal.
 *
 * <p>PowerPoint sait afficher uniquement le caractere symbole avec la police
 * "sym" (Wingdings) et le reste avec la police normale heritee. Apache POI,
 * lui, semble appliquer la police "sym" a la totalite du texte du run des
 * qu'aucun {@code <a:latin>} explicite n'est present - et comme une police
 * comme Wingdings fait correspondre chaque lettre latine a un pictogramme,
 * tout le texte concerne se retrouve rendu comme une suite de symboles
 * illisibles au lieu du texte attendu.
 *
 * <p><b>Trois variantes du meme motif ont ete observees dans des fichiers
 * reels</b> :
 * <ol>
 *   <li>un run UNIQUE melange caractere symbole et texte normal
 *       (ex. {@code " modifier les conditions d'acces..."}) ;</li>
 *   <li>la police sym reste active sur PLUSIEURS runs consecutifs jusqu'a la
 *       fin du paragraphe, y compris des runs qui ne contiennent aucun
 *       caractere symbole du tout (ex. un mot isole dans son propre run par
 *       la verification orthographique de PowerPoint, mais qui herite quand
 *       meme du {@code <a:sym>} pose au moment de la frappe du caractere
 *       special juste avant). Une premiere version de ce correctif ne
 *       traitait que le cas (1) et seulement quand le run mixte etait le
 *       DERNIER run du paragraphe - trop restrictif : elle laissait ce cas
 *       (2) totalement intact ;</li>
 *   <li><b>slide 25 de "Mes Evenements Emploi et Prestation - Doc vision
 *       1.0.pptx"</b> : un run isole ne contenant AUCUN caractere symbole
 *       (juste {@code " p"}, un espace et la lettre latine ordinaire "p" -
 *       coupure malheureuse d'un mot en debut de frappe d'un caractere
 *       special, jamais finalisee) porte quand meme un {@code <a:sym
 *       typeface="Wingdings">}, et est immediatement SUIVI d'un run tout a
 *       fait normal, SANS aucune police sym ({@code "oint d'entree commun
 *       pour les conseillers"} - la suite du mot "point", coupe juste apres
 *       le "p"). Apache POI applique donc Wingdings a la lettre "p" isolee,
 *       qui devient un carre (glyphe absent), alors qu'une ligne quasiment
 *       identique juste en dessous ({@code "Point d'entree commun pour les
 *       candidats"}) s'affiche correctement : sur celle-ci, le run sym ne
 *       contient QUE l'espace (texte {@code " "}), et le "P" (majuscule) fait
 *       entierement partie du run normal suivant - la coupure du mot a ete
 *       faite au bon endroit par l'auteur, contrairement a la ligne du
 *       dessus.</li>
 * </ol>
 *
 * <p><b>Correctif</b> : avant le rendu, on repere pour chaque paragraphe le
 * premier run "problematique" (police sym posee sur un texte qui n'est pas
 * <em>entierement</em> compose de caracteres de la zone d'usage prive Unicode
 * U+E000-U+F8FF, la zone ou vivent les caracteres symboles Wingdings/Webdings/
 * Symbol dans un fichier Office - le cas (3) ci-dessus, sans le moindre
 * caractere symbole du tout, est un cas degenere mais couvert par cette meme
 * definition). On reconstruit alors la "zone" entiere, de ce premier run
 * problematique jusqu'a la toute fin du paragraphe (que les runs suivants
 * portent ou non une police sym - voir "Limite assumee" ci-dessous pour la
 * seule vraie limite restante) : chaque run de la zone est redecoupe en
 * segments symbole/texte normal (un run sans aucune police sym, cas (3),
 * produit un seul segment de texte normal, recopie tel quel sans police
 * symbole a appliquer), le premier segment reutilise le run existant, et tous
 * les segments suivants sont ajoutes a la suite via
 * {@link XSLFTextParagraph#addNewTextRun()} - la methode publique de POI pour
 * ajouter un run, seule garantie de rester coherente avec l'etat interne
 * (cache) que POI maintient pour un paragraphe. Les runs d'origine de la zone
 * autres que le premier sont vides (texte remplace par {@code ""}) : leur
 * contenu vit desormais dans les nouveaux runs ajoutes, dans le bon ordre -
 * un run vide ne rend visuellement rien, donc le resultat est correct meme
 * s'il reste quelques {@code <a:r>} vides dans le XML.
 *
 * <p><b>Limite assumee</b> : si un run de la zone n'est pas un {@code
 * CTRegularTextRun} ordinaire (ex. un saut de ligne ou un champ automatique -
 * motif non rencontre a ce jour), on renonce et on laisse tout le paragraphe
 * intact plutot que de risquer de corrompre l'ordre du texte ({@code
 * addNewTextRun()} ne peut qu'ajouter en fin de paragraphe, donc une
 * insertion "au milieu" n'est pas possible sans ce risque, d'ou la necessite
 * de reconstruire TOUTE la zone jusqu'a la fin plutot que seulement les runs
 * effectivement fautifs).
 *
 * <p>Une premiere version de ce correctif manipulait directement l'arbre XML
 * sous-jacent ({@code CTTextParagraph.insertNewR()}/{@code removeR()}) : les
 * tests ont revele que cette approche desynchronise l'etat interne que POI
 * maintient pour les runs d'un paragraphe - d'ou le choix de repasser
 * exclusivement par l'API publique de POI pour toute creation de run.
 */
public final class SymbolFontRunFixer {

    private static final Logger LOG = LoggerFactory.getLogger(SymbolFontRunFixer.class);

    private SymbolFontRunFixer() {
    }

    /**
     * Corrige, in-place, les runs "mixtes" (police symbole appliquee a du
     * texte normal) du slide. A appeler avant {@code slide.draw(graphics)}.
     *
     * @return le nombre de runs effectivement touches par une correction.
     */
    public static int fixMixedSymbolRuns(XSLFSlide slide) {
        int fixed = 0;
        for (XSLFShape shape : collectTextShapes(slide.getShapes())) {
            XSLFTextShape ts = (XSLFTextShape) shape;
            for (XSLFTextParagraph para : ts.getTextParagraphs()) {
                fixed += fixParagraph(para);
            }
        }
        return fixed;
    }

    private static int fixParagraph(XSLFTextParagraph para) {
        // Instantane pris avant toute mutation.
        List<XSLFTextRun> originalRuns = new ArrayList<>(para.getTextRuns());

        int firstProblemIdx = -1;
        for (int i = 0; i < originalRuns.size(); i++) {
            if (isProblematicSymRun(originalRuns.get(i))) {
                firstProblemIdx = i;
                break;
            }
        }
        if (firstProblemIdx == -1) {
            return 0;
        }

        // La zone couvre toujours jusqu'a la toute fin du paragraphe (voir
        // Javadoc de la classe, "addNewTextRun() ne peut qu'ajouter en fin de
        // paragraphe") - qu'un run de la zone porte ou non une police sym
        // (cas (3) de la Javadoc : un run normal, sans sym, qui suit un run
        // problematique). On renonce uniquement si un run de la zone n'est
        // pas un CTRegularTextRun ordinaire, seul cas non reconstructible en
        // toute securite.
        for (int i = firstProblemIdx; i < originalRuns.size(); i++) {
            if (asRegularTextRun(originalRuns.get(i)) == null) {
                if (LOG.isDebugEnabled()) {
                    LOG.debug("run non reconstructible (pas un CTRegularTextRun ordinaire) dans la zone a corriger, "
                            + "correction non tentee");
                }
                return 0;
            }
        }

        XSLFTextRun firstRun = originalRuns.get(firstProblemIdx);
        CTRegularTextRun ctFirst = asRegularTextRun(firstRun);
        String firstTypeface = ctFirst.getRPr().getSym().getTypeface();
        Double firstFontSize = firstRun.getFontSize();
        Boolean firstBold = firstRun.isBold();
        Boolean firstItalic = firstRun.isItalic();
        // Capturees ICI, avant toute mutation de firstRun (setText/clearSym plus bas) : un
        // nouveau run cree par XSLFTextParagraph#addNewTextRun() ne reprend NI la police
        // latine NI la couleur du run precedent (contrairement a la taille/gras/italique deja
        // reportes ci-dessous) - sans report explicite, chaque nouveau run bascule sur les
        // valeurs par defaut heritees (police et couleur du masque/theme), ce qui denature
        // visiblement le texte reconstruit (constate en pratique : slide 7 de "Mes Evenements
        // Emploi et Prestation - Doc vision 1.0.pptx", segment de texte normal suivant la fleche
        // Wingdings d'un run mixte affiche en gras et en bleu marine au lieu du gris attendu).
        //
        // getFontFamily(FontGroup.LATIN) - PAS l'overload sans argument - est indispensable
        // ici : ce dernier determine le groupe de police via FontGroup.getFontGroupFirst(texte
        // ACTUEL du run), et a cet instant precis firstRun porte encore son texte D'ORIGINE
        // (symbole + texte normal), qui COMMENCE justement par le caractere symbole (c'est le
        // motif meme qui declenche ce correctif) - l'overload sans argument renverrait donc la
        // police "sym" (Wingdings) au lieu de la police latine attendue, et
        // newRun.setFontFamily(...) plus bas l'appliquerait comme police LATINE du nouveau run
        // (nouveau bug constate en testant ce correctif : texte reconstruit affiche avec les
        // glyphes de la police de substitution systeme pour "Wingdings" au lieu du texte normal).
        String firstFontFamily = firstRun.getFontFamily(FontGroup.LATIN);
        PaintStyle firstFontColor = firstRun.getFontColor();
        List<String> firstSegments = splitSymbolSegments(firstRun.getRawText());

        // Le run existant est reduit a son premier segment - seul segment dont
        // on reutilise le run d'origine.
        firstRun.setText(firstSegments.get(0));
        if (!isSymbolChar(firstSegments.get(0).charAt(0))) {
            // Ce premier segment n'est en fait pas un symbole (run "pur texte"
            // mais errone-ment marque sym) : on retire le sym errone.
            clearSym(firstRun);
        }

        List<String> pendingText = new ArrayList<>();
        List<String> pendingTypeface = new ArrayList<>(); // null = pas de police symbole pour ce segment
        List<Double> pendingFontSize = new ArrayList<>();
        List<Boolean> pendingBold = new ArrayList<>();
        List<Boolean> pendingItalic = new ArrayList<>();
        List<String> pendingFontFamily = new ArrayList<>();
        List<PaintStyle> pendingFontColor = new ArrayList<>();

        for (int s = 1; s < firstSegments.size(); s++) {
            String seg = firstSegments.get(s);
            pendingText.add(seg);
            pendingTypeface.add(isSymbolChar(seg.charAt(0)) ? firstTypeface : null);
            pendingFontSize.add(firstFontSize);
            pendingBold.add(firstBold);
            pendingItalic.add(firstItalic);
            pendingFontFamily.add(firstFontFamily);
            pendingFontColor.add(firstFontColor);
        }

        for (int i = firstProblemIdx + 1; i < originalRuns.size(); i++) {
            XSLFTextRun run = originalRuns.get(i);
            CTRegularTextRun ctRun = asRegularTextRun(run);
            CTTextCharacterProperties rPr = ctRun.getRPr();
            // Cas (3) de la Javadoc : ce run peut tout a fait n'avoir aucune police
            // sym (run normal qui suit, dans la meme zone, un run problematique) -
            // traite alors comme un simple segment de texte normal, sans police
            // symbole a appliquer.
            boolean hasSym = rPr != null && rPr.isSetSym();
            String typeface = hasSym ? rPr.getSym().getTypeface() : null;
            Double fontSize = run.getFontSize();
            Boolean bold = run.isBold();
            Boolean italic = run.isItalic();
            // Capturees, comme firstFontFamily/firstFontColor plus haut (meme raison pour
            // getFontFamily(FontGroup.LATIN) plutot que l'overload sans argument - ce run peut,
            // lui aussi, avoir un texte qui commence par un caractere symbole), AVANT le
            // run.setText("") qui vide ce run juste en dessous.
            String fontFamily = run.getFontFamily(FontGroup.LATIN);
            PaintStyle fontColor = run.getFontColor();
            for (String seg : splitSymbolSegments(run.getRawText())) {
                pendingText.add(seg);
                pendingTypeface.add(isSymbolChar(seg.charAt(0)) ? typeface : null);
                pendingFontSize.add(fontSize);
                pendingBold.add(bold);
                pendingItalic.add(italic);
                pendingFontFamily.add(fontFamily);
                pendingFontColor.add(fontColor);
            }
            // Ce run d'origine est vide - son contenu vit desormais dans les
            // nouveaux runs ajoutes plus bas, dans le bon ordre. Un run vide
            // ne rend visuellement rien.
            run.setText("");
        }

        for (int s = 0; s < pendingText.size(); s++) {
            XSLFTextRun newRun = para.addNewTextRun();
            newRun.setText(pendingText.get(s));
            if (pendingFontSize.get(s) != null) {
                newRun.setFontSize(pendingFontSize.get(s));
            }
            if (pendingBold.get(s) != null) {
                newRun.setBold(pendingBold.get(s));
            }
            if (pendingItalic.get(s) != null) {
                newRun.setItalic(pendingItalic.get(s));
            }
            // Sans ce report explicite, un nouveau run demarre sans police latine ni
            // couleur propres et retombe sur les valeurs par defaut heritees du
            // masque/theme - voir Javadoc de firstFontFamily/firstFontColor plus haut.
            if (pendingFontFamily.get(s) != null) {
                newRun.setFontFamily(pendingFontFamily.get(s));
            }
            if (pendingFontColor.get(s) != null) {
                newRun.setFontColor(pendingFontColor.get(s));
            }
            String typeface = pendingTypeface.get(s);
            if (typeface != null) {
                applySymTypeface(newRun, typeface);
            }
        }

        int touched = originalRuns.size() - firstProblemIdx;
        if (LOG.isDebugEnabled()) {
            LOG.debug("paragraphe corrige : {} run(s) symbole reconstruit(s) en {} segment(s)",
                    touched, 1 + pendingText.size());
        }
        return touched;
    }

    /**
     * Un run est "problematique" s'il porte une police sym mais que son texte
     * n'est pas <em>entierement</em> compose de caracteres symboles - que ce
     * soit un melange (symbole + texte normal dans le meme run) ou meme du
     * texte normal pur (aucun caractere symbole du tout, mais police sym
     * quand meme posee).
     */
    private static boolean isProblematicSymRun(XSLFTextRun run) {
        CTRegularTextRun ctRun = asRegularTextRun(run);
        if (ctRun == null) {
            return false;
        }
        CTTextCharacterProperties rPr = ctRun.getRPr();
        String text = run.getRawText();
        if (rPr == null || !rPr.isSetSym() || text == null || text.isEmpty()) {
            return false;
        }
        List<String> segments = splitSymbolSegments(text);
        boolean pureSymbolRun = segments.size() == 1 && isSymbolChar(segments.get(0).charAt(0));
        return !pureSymbolRun;
    }

    private static void clearSym(XSLFTextRun run) {
        CTRegularTextRun ctRun = asRegularTextRun(run);
        if (ctRun != null && ctRun.isSetRPr() && ctRun.getRPr().isSetSym()) {
            ctRun.getRPr().unsetSym();
        }
    }

    private static void applySymTypeface(XSLFTextRun run, String typeface) {
        CTRegularTextRun ctRun = asRegularTextRun(run);
        if (ctRun == null) {
            return;
        }
        CTTextCharacterProperties rPr = ctRun.isSetRPr() ? ctRun.getRPr() : ctRun.addNewRPr();
        CTTextFont sym = rPr.addNewSym();
        sym.setTypeface(typeface);
    }

    private static CTRegularTextRun asRegularTextRun(XSLFTextRun run) {
        Object xmlObject = run.getXmlObject();
        return (xmlObject instanceof CTRegularTextRun) ? (CTRegularTextRun) xmlObject : null;
    }

    /** Zone d'usage prive Unicode (U+E000-U+F8FF) - c'est la que vivent les caracteres des polices symboles Office (Wingdings, Webdings, Symbol...). */
    private static boolean isSymbolChar(char c) {
        return c >= 0xE000 && c <= 0xF8FF;
    }

    /** Decoupe le texte en segments consecutifs de meme "nature" (symbole vs texte normal). */
    private static List<String> splitSymbolSegments(String text) {
        List<String> segments = new ArrayList<>();
        if (text.isEmpty()) {
            segments.add(text);
            return segments;
        }
        StringBuilder current = new StringBuilder();
        Boolean currentIsSymbol = null;
        for (int i = 0; i < text.length(); i++) {
            char c = text.charAt(i);
            boolean symbol = isSymbolChar(c);
            if (currentIsSymbol != null && symbol != currentIsSymbol) {
                segments.add(current.toString());
                current = new StringBuilder();
            }
            current.append(c);
            currentIsSymbol = symbol;
        }
        if (current.length() > 0) {
            segments.add(current.toString());
        }
        return segments;
    }

    /**
     * Parcourt les formes du slide, y compris a l'interieur des groupes ET des
     * cellules de tableau, et ne garde que celles porteuses de texte.
     *
     * <p>Bug corrige : un {@code XSLFTable} n'est ni un {@code XSLFGroupShape}
     * ni un {@code XSLFTextShape} - c'est une forme a part, dont les cellules
     * ({@code XSLFTableCell}, qui EST un {@code XSLFTextShape}) ne sont
     * accessibles que via {@code table.getRows()} / {@code row.getCells()}.
     * Sans ce cas particulier, tout texte dans un tableau (constate sur un
     * fichier reel) etait completement ignore par le correctif.
     */
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
