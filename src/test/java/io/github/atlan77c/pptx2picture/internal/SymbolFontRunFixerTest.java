package io.github.atlan77c.pptx2picture.internal;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTable;
import org.apache.poi.xslf.usermodel.XSLFTableCell;
import org.apache.poi.xslf.usermodel.XSLFTableRow;
import org.apache.poi.xslf.usermodel.XSLFTextBox;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.drawingml.x2006.main.CTRegularTextRun;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextCharacterProperties;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextFont;

import java.awt.geom.Rectangle2D;
import java.io.IOException;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * Reproduit le motif observe dans un fichier reel : un run unique portant
 * {@code <a:sym typeface="Wingdings">} sans {@code <a:latin>}, dont le texte
 * melange un caractere de la zone d'usage prive Unicode (celui destine a
 * Wingdings) et du texte normal accole juste apres. Les fixtures construisent
 * ce motif directement via l'API XML sous-jacente de POI (pas d'API haut
 * niveau publique pour poser un {@code <a:sym>}), exactement comme le ferait
 * un fichier .pptx exporte par PowerPoint.
 */
class SymbolFontRunFixerTest {

    /** U+F0E8 = fleche Wingdings dans la zone d'usage prive Unicode - exactement le caractere observe dans un fichier reel. */
    private static final String WINGDINGS_ARROW = "";

    private XMLSlideShow ppt;
    private XSLFSlide slide;

    @BeforeEach
    void setUp() {
        ppt = new XMLSlideShow();
        slide = ppt.createSlide();
    }

    @AfterEach
    void tearDown() throws IOException {
        ppt.close();
    }

    private XSLFTextBox textBox(String text) {
        XSLFTextBox box = slide.createTextBox();
        box.setAnchor(new Rectangle2D.Double(10, 10, 300, 100));
        box.setText(text);
        return box;
    }

    /** Pose <a:sym typeface="Wingdings"/> sur le run, sans <a:latin> - reproduit le motif du fichier reel. */
    private void applyMixedSymbolFont(XSLFTextRun run) {
        CTRegularTextRun ctRun = (CTRegularTextRun) run.getXmlObject();
        CTTextCharacterProperties rPr = ctRun.isSetRPr() ? ctRun.getRPr() : ctRun.addNewRPr();
        CTTextFont sym = rPr.addNewSym();
        sym.setTypeface("Wingdings");
    }

    @Test
    void fixMixedSymbolRuns_splitsRunIntoSymbolSegmentAndPlainTextSegment() {
        XSLFTextBox box = textBox(WINGDINGS_ARROW + " modifier les conditions");
        XSLFTextRun run = box.getTextParagraphs().get(0).getTextRuns().get(0);
        applyMixedSymbolFont(run);

        int fixed = SymbolFontRunFixer.fixMixedSymbolRuns(slide);
        assertEquals(1, fixed);

        List<XSLFTextRun> runsAfter = box.getTextParagraphs().get(0).getTextRuns();
        assertEquals(2, runsAfter.size(), "le run mixte doit avoir ete scinde en 2 : symbole puis texte normal");

        XSLFTextRun symbolRun = runsAfter.get(0);
        XSLFTextRun textRun = runsAfter.get(1);

        assertEquals(WINGDINGS_ARROW, symbolRun.getRawText());
        assertEquals(" modifier les conditions", textRun.getRawText());

        CTRegularTextRun ctSymbol = (CTRegularTextRun) symbolRun.getXmlObject();
        assertTrue(ctSymbol.getRPr().isSetSym(), "le segment symbole doit conserver sa police Wingdings");

        CTRegularTextRun ctText = (CTRegularTextRun) textRun.getXmlObject();
        assertFalse(ctText.getRPr().isSetSym(), "le segment de texte normal ne doit plus avoir de police symbole");
    }

    @Test
    void fixMixedSymbolRuns_leavesPureSymbolRunUntouched() {
        // Run ne contenant QUE le caractere symbole, sans texte accole : rien a scinder.
        XSLFTextBox box = textBox(WINGDINGS_ARROW);
        XSLFTextRun run = box.getTextParagraphs().get(0).getTextRuns().get(0);
        applyMixedSymbolFont(run);

        int fixed = SymbolFontRunFixer.fixMixedSymbolRuns(slide);
        assertEquals(0, fixed);

        List<XSLFTextRun> runsAfter = box.getTextParagraphs().get(0).getTextRuns();
        assertEquals(1, runsAfter.size());
        CTRegularTextRun ctRun = (CTRegularTextRun) runsAfter.get(0).getXmlObject();
        assertTrue(ctRun.getRPr().isSetSym(), "un run purement symbole ne doit pas etre modifie");
    }

    @Test
    void fixMixedSymbolRuns_leavesNormalRunsUntouched() {
        XSLFTextBox box = textBox("Texte tout a fait normal, sans aucun caractere symbole");
        int fixed = SymbolFontRunFixer.fixMixedSymbolRuns(slide);
        assertEquals(0, fixed);
        assertEquals(1, box.getTextParagraphs().get(0).getTextRuns().size());
    }

    @Test
    void fixMixedSymbolRuns_fixesSymbolTaintedZoneSpanningMultipleTrailingRuns() {
        // Reproduit exactement un motif observe sur un fichier reel : la police
        // sym reste active sur PLUSIEURS runs consecutifs jusqu'a la fin du
        // paragraphe, y compris des runs qui ne contiennent aucun caractere
        // symbole du tout (ex. un mot isole scinde dans son propre run par la
        // verification orthographique de PowerPoint, mais tague sym quand meme).
        XSLFTextBox box = slide.createTextBox();
        box.setAnchor(new Rectangle2D.Double(10, 10, 400, 150));
        XSLFTextParagraph para = box.addNewTextParagraph();

        addPlainRun(para, "Un texte assez long qui se termine par un mot isole : ");
        addPlainRun(para, "Contoso");
        addPlainRun(para, " ");
        XSLFTextRun mixed = addPlainRun(para, WINGDINGS_ARROW + " voir la proposition : distinguer l'entite (");
        applyMixedSymbolFont(mixed);
        XSLFTextRun tail1 = addPlainRun(para, "Contoso");
        applyMixedSymbolFont(tail1);
        XSLFTextRun tail2 = addPlainRun(para, " ou l'autre) ?");
        applyMixedSymbolFont(tail2);

        String fullTextBefore = concatText(para);

        int fixed = SymbolFontRunFixer.fixMixedSymbolRuns(slide);
        assertEquals(3, fixed, "les 3 runs de la zone symbole (le run mixte + les 2 runs texte pur mal tagues) doivent etre touches");

        // Le texte visible complet doit rester identique, dans le meme ordre.
        assertEquals(fullTextBefore, concatText(para));

        // Plus aucun run ne doit melanger police symbole et texte non-symbole,
        // ni porter une police symbole sur du texte purement normal.
        for (XSLFTextRun r : para.getTextRuns()) {
            String t = r.getRawText();
            if (t == null || t.isEmpty()) {
                continue;
            }
            CTRegularTextRun ct = (CTRegularTextRun) r.getXmlObject();
            boolean hasSym = ct.getRPr() != null && ct.getRPr().isSetSym();
            boolean allSymbolChars = t.chars().allMatch(c -> c >= 0xE000 && c <= 0xF8FF);
            if (hasSym) {
                assertTrue(allSymbolChars, "un run encore marque police symbole ne doit contenir QUE des caracteres symboles : " + t);
            }
        }
    }

    @Test
    void fixMixedSymbolRuns_fixesSymTaggedRunWithoutAnySymbolCharFollowedByNormalRun() {
        // Reproduit exactement le motif du slide 25 de "Mes Evenements Emploi et
        // Prestation - Doc vision 1.0.pptx" (voir Javadoc de la classe, cas (3)) :
        // un run isole " p" (espace + lettre latine ordinaire, AUCUN caractere
        // symbole) porte quand meme <a:sym typeface="Wingdings">, et est
        // immediatement suivi d'un run tout a fait normal, sans la moindre
        // police sym - contrairement au cas (2), le run sym n'est PAS suivi
        // d'autres runs sym : c'etait le motif que l'ancienne version de ce
        // correctif abandonnait entierement (zone non contigue jusqu'a la fin).
        XSLFTextBox box = slide.createTextBox();
        box.setAnchor(new Rectangle2D.Double(10, 10, 400, 100));
        XSLFTextParagraph para = box.addNewTextParagraph();

        addPlainRun(para, "Unification du persona « conseiller »  ");
        XSLFTextRun brokenP = addPlainRun(para, " p");
        applyMixedSymbolFont(brokenP);
        addPlainRun(para, "oint d’entrée commun pour les conseillers");

        String fullTextBefore = concatText(para);

        int fixed = SymbolFontRunFixer.fixMixedSymbolRuns(slide);
        assertEquals(2, fixed, "les 2 runs de la zone (le 'p' isole mal tague + le run normal qui suit) doivent etre touches");

        // Le texte visible complet doit rester identique, dans le meme ordre -
        // en particulier, le "p" ne doit pas migrer apres "oint d'entree...".
        assertEquals(fullTextBefore, concatText(para));

        // Plus aucun run ne doit porter de police symbole sur du texte non-symbole.
        for (XSLFTextRun r : para.getTextRuns()) {
            String t = r.getRawText();
            if (t == null || t.isEmpty()) {
                continue;
            }
            CTRegularTextRun ct = (CTRegularTextRun) r.getXmlObject();
            boolean hasSym = ct.getRPr() != null && ct.getRPr().isSetSym();
            boolean allSymbolChars = t.chars().allMatch(c -> c >= 0xE000 && c <= 0xF8FF);
            assertFalse(hasSym && !allSymbolChars, "run encore marque police symbole sur du texte non-symbole : " + t);
        }
    }

    @Test
    void fixMixedSymbolRuns_leavesVisuallyCorrectSpaceOnlySymRunHarmless() {
        // Motif "correct" observe juste en dessous du bug, sur le meme slide 25 :
        // le run sym ne contient QUE l'espace (aucun caractere visible, donc
        // policesym ou pas ne change rien visuellement), et le mot suivant
        // ("Point", majuscule) est entierement dans le run normal qui suit. Ce
        // paragraphe est techniquement repere comme "a corriger" (l'espace seul
        // n'est pas un caractere de la zone symbole - voir isProblematicSymRun),
        // mais le resultat doit rester visuellement identique : aucune perte ni
        // deplacement de texte.
        XSLFTextBox box = slide.createTextBox();
        box.setAnchor(new Rectangle2D.Double(10, 10, 400, 100));
        XSLFTextParagraph para = box.addNewTextParagraph();

        XSLFTextRun spaceOnly = addPlainRun(para, " ");
        applyMixedSymbolFont(spaceOnly);
        addPlainRun(para, "Point d’entrée commun pour les candidats");

        String fullTextBefore = concatText(para);

        SymbolFontRunFixer.fixMixedSymbolRuns(slide);

        assertEquals(fullTextBefore, concatText(para), "le texte visible doit rester strictement identique");
        for (XSLFTextRun r : para.getTextRuns()) {
            String t = r.getRawText();
            if (t == null || t.isEmpty()) {
                continue;
            }
            CTRegularTextRun ct = (CTRegularTextRun) r.getXmlObject();
            boolean hasSym = ct.getRPr() != null && ct.getRPr().isSetSym();
            boolean allSymbolChars = t.chars().allMatch(c -> c >= 0xE000 && c <= 0xF8FF);
            assertFalse(hasSym && !allSymbolChars, "run encore marque police symbole sur du texte non-symbole : " + t);
        }
    }

    private XSLFTextRun addPlainRun(XSLFTextParagraph para, String text) {
        XSLFTextRun run = para.addNewTextRun();
        run.setText(text);
        return run;
    }

    private String concatText(XSLFTextParagraph para) {
        StringBuilder sb = new StringBuilder();
        for (XSLFTextRun r : para.getTextRuns()) {
            if (r.getRawText() != null) {
                sb.append(r.getRawText());
            }
        }
        return sb.toString();
    }

    @Test
    void fixMixedSymbolRuns_fixesRunsInsideTableCells() {
        // Bug corrige : un XSLFTable n'est ni un XSLFGroupShape ni un
        // XSLFTextShape - c'est une forme a part, dont les cellules ne sont
        // accessibles que via table.getRows()/row.getCells(). Sans le cas
        // particulier ajoute dans collectTextShapes(), tout texte dans un
        // tableau (constate sur un tableau d'un fichier reel) etait
        // completement ignore par le correctif, meme si le run etait par
        // ailleurs un cas simple deja gere (mixte, dernier run du paragraphe).
        XSLFTable table = slide.createTable();
        XSLFTableRow row = table.addRow();
        XSLFTableCell cell = row.addCell();
        XSLFTextParagraph para = cell.addNewTextParagraph();
        XSLFTextRun run = para.addNewTextRun();
        run.setText(WINGDINGS_ARROW + " modifie par relecture manuelle");
        applyMixedSymbolFont(run);

        int fixed = SymbolFontRunFixer.fixMixedSymbolRuns(slide);
        assertEquals(1, fixed, "le run mixte dans la cellule de tableau doit etre corrige");

        List<XSLFTextRun> runsAfter = para.getTextRuns();
        assertEquals(2, runsAfter.size(), "le run mixte de la cellule doit avoir ete scinde en 2");
        assertEquals(WINGDINGS_ARROW, runsAfter.get(0).getRawText());
        assertEquals(" modifie par relecture manuelle", runsAfter.get(1).getRawText());

        CTRegularTextRun ctText = (CTRegularTextRun) runsAfter.get(1).getXmlObject();
        assertFalse(ctText.getRPr() != null && ctText.getRPr().isSetSym(),
                "le segment de texte normal de la cellule ne doit plus avoir de police symbole");
    }

    @Test
    void fixMixedSymbolRuns_splitsEvenWhenLatinAlreadySet() {
        // Motif hypothetique (non observe dans un fichier reel a ce jour) : le run a
        // deja un <a:latin> explicite EN PLUS du sym. La detection ne se base pas sur
        // l'absence de latin (uniquement sur le melange symbole/texte normal dans un
        // run sym), donc la scission a lieu quand meme - le segment symbole garde tout
        // son rPr d'origine (latin + sym inchanges), seul le nouveau segment de texte
        // normal n'a pas de police symbole.
        XSLFTextBox box = textBox(WINGDINGS_ARROW + " texte normal");
        XSLFTextRun run = box.getTextParagraphs().get(0).getTextRuns().get(0);
        run.setFontFamily("Custom Sans");
        applyMixedSymbolFont(run);

        int fixed = SymbolFontRunFixer.fixMixedSymbolRuns(slide);
        assertEquals(1, fixed);

        List<XSLFTextRun> runsAfter = box.getTextParagraphs().get(0).getTextRuns();
        assertEquals(2, runsAfter.size());
        CTRegularTextRun ctText = (CTRegularTextRun) runsAfter.get(1).getXmlObject();
        assertFalse(ctText.getRPr() != null && ctText.getRPr().isSetSym(),
                "le segment de texte normal ne doit pas avoir de police symbole, meme si l'original avait aussi un latin explicite");
    }
}
