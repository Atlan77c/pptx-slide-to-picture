package io.github.atlan77c.pptx2picture.internal;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextBox;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.io.IOException;

import static org.junit.jupiter.api.Assertions.assertEquals;

/**
 * Reproduit le motif observe sur un fichier reel (slide 10 de
 * "02.CGSI 2023 04 11 - SI Fraude_DIANE-CGSI_11avril2023_v1.0.pptx") : un
 * paragraphe a interligne en pourcentage contenant un run entierement blanc
 * (un espace) a une taille de police bien plus grande que le texte visible du
 * meme paragraphe - Apache POI mesure la hauteur de cette ligne a partir de
 * TOUS ses runs, blancs compris, ce qui gonfle la ligne et pousse tout le
 * contenu suivant plus bas que chez PowerPoint.
 */
class OversizedWhitespaceRunFixerTest {

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

    @Test
    void fixOversizedWhitespaceRuns_shrinksOversizedBlankRun_whenPercentageLineSpacing() {
        XSLFTextBox box = slide.createTextBox();
        box.setText("817 ");
        XSLFTextParagraph para = box.getTextParagraphs().get(0);
        para.setLineSpacing(100.0); // positif = pourcentage, comme <a:spcPct val="100000"/>
        para.getTextRuns().get(0).setFontSize(20.0);
        addRun(para, "signalements crees", 14.0);
        addRun(para, " ", 48.0); // motif reel : espace isole a 48pt

        int fixed = OversizedWhitespaceRunFixer.fixOversizedWhitespaceRuns(slide);

        assertEquals(1, fixed);
        assertEquals(20.0, para.getTextRuns().get(2).getFontSize());
        // le texte visible n'est jamais touche
        assertEquals("817 ", para.getTextRuns().get(0).getRawText());
        assertEquals(20.0, para.getTextRuns().get(0).getFontSize());
        assertEquals("signalements crees", para.getTextRuns().get(1).getRawText());
        assertEquals(14.0, para.getTextRuns().get(1).getFontSize());
    }

    @Test
    void fixOversizedWhitespaceRuns_leavesUntouched_whenLineSpacingIsAbsolute() {
        XSLFTextBox box = slide.createTextBox();
        box.setText("34 ");
        XSLFTextParagraph para = box.getTextParagraphs().get(0);
        para.setLineSpacing(-16.8); // negatif = points absolus, comme <a:spcPts val="1680"/>
        para.getTextRuns().get(0).setFontSize(20.0);
        addRun(para, "auditeurs fraude", 14.0);
        addRun(para, " ", 48.0);

        int fixed = OversizedWhitespaceRunFixer.fixOversizedWhitespaceRuns(slide);

        assertEquals(0, fixed, "interligne absolu : deja insensible aux metriques de police, rien a corriger");
        assertEquals(48.0, para.getTextRuns().get(2).getFontSize());
    }

    @Test
    void fixOversizedWhitespaceRuns_leavesUntouched_whenBlankRunNotLargerThanVisibleText() {
        XSLFTextBox box = slide.createTextBox();
        box.setText("Texte normal");
        XSLFTextParagraph para = box.getTextParagraphs().get(0);
        para.setLineSpacing(100.0);
        para.getTextRuns().get(0).setFontSize(20.0);
        addRun(para, " ", 14.0); // plus petit que le texte visible : pas anormal

        int fixed = OversizedWhitespaceRunFixer.fixOversizedWhitespaceRuns(slide);

        assertEquals(0, fixed);
        assertEquals(14.0, para.getTextRuns().get(1).getFontSize());
    }

    @Test
    void fixOversizedWhitespaceRuns_leavesFullyBlankParagraphUntouched() {
        // Paragraphe-espaceur vide (aucun texte visible) : pas de reference a laquelle
        // comparer, ne doit jamais etre touche meme avec une taille inhabituelle.
        XSLFTextBox box = slide.createTextBox();
        box.setText(" ");
        XSLFTextParagraph para = box.getTextParagraphs().get(0);
        para.setLineSpacing(100.0);
        para.getTextRuns().get(0).setFontSize(48.0);

        int fixed = OversizedWhitespaceRunFixer.fixOversizedWhitespaceRuns(slide);

        assertEquals(0, fixed, "paragraphe sans texte visible : aucune reference, ne pas toucher");
        assertEquals(48.0, para.getTextRuns().get(0).getFontSize());
    }

    @Test
    void fixOversizedWhitespaceRuns_treatsNullLineSpacingAsPercentageBased() {
        // Interligne non declare (ni local ni herite) : se comporte comme un interligne
        // en pourcentage cote POI (base sur les metriques mesurees), donc concerne aussi.
        XSLFTextBox box = slide.createTextBox();
        box.setText("Texte");
        XSLFTextParagraph para = box.getTextParagraphs().get(0);
        para.getTextRuns().get(0).setFontSize(20.0);
        addRun(para, " ", 48.0);

        int fixed = OversizedWhitespaceRunFixer.fixOversizedWhitespaceRuns(slide);

        assertEquals(1, fixed);
        assertEquals(20.0, para.getTextRuns().get(1).getFontSize());
    }

    private static void addRun(XSLFTextParagraph para, String text, double fontSize) {
        XSLFTextRun run = para.addNewTextRun();
        run.setText(text);
        run.setFontSize(fontSize);
    }
}
