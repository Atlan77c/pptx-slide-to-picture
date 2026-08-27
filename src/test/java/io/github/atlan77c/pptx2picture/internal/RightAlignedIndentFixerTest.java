package io.github.atlan77c.pptx2picture.internal;

import org.apache.poi.sl.usermodel.TextParagraph.TextAlign;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextBox;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.io.IOException;

import static org.junit.jupiter.api.Assertions.assertEquals;

/**
 * Reproduit le motif observe sur deux fichiers reels : un paragraphe aligne a
 * droite (ex. "P.3" d'un sommaire) avec un {@code marL}/{@code indent} non
 * nul (motif de liste a puces avec indentation suspendue, {@code marL=22.5pt
 * indent=-22.5pt} dans le premier fichier) - Apache POI ne rend pas
 * correctement cette combinaison, le texte deborde entierement de sa boite
 * vers la droite et se superpose au debut de la forme voisine.
 */
class RightAlignedIndentFixerTest {

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
    void fixInheritedIndent_clearsMarginAndIndent_forRightAlignedParagraph() {
        XSLFTextBox box = slide.createTextBox();
        box.setText("P.3");
        XSLFTextParagraph para = box.getTextParagraphs().get(0);
        para.setTextAlign(TextAlign.RIGHT);
        para.setLeftMargin(22.5);
        para.setIndent(-22.5);

        int fixed = RightAlignedIndentFixer.fixInheritedIndent(slide);

        assertEquals(1, fixed);
        assertEquals(0.0, para.getLeftMargin());
        assertEquals(0.0, para.getIndent());
    }

    @Test
    void fixInheritedIndent_clearsMarginAndIndent_forCenterAlignedParagraph() {
        XSLFTextBox box = slide.createTextBox();
        box.setText("Titre centre");
        XSLFTextParagraph para = box.getTextParagraphs().get(0);
        para.setTextAlign(TextAlign.CENTER);
        para.setLeftMargin(14.0);
        para.setIndent(-14.0);

        int fixed = RightAlignedIndentFixer.fixInheritedIndent(slide);

        assertEquals(1, fixed);
        assertEquals(0.0, para.getLeftMargin());
        assertEquals(0.0, para.getIndent());
    }

    @Test
    void fixInheritedIndent_leavesLeftAlignedParagraphsUntouched() {
        // Cas normal d'une liste a puces : marL/indent ont un sens pour de l'aligne a
        // gauche, ne doit jamais etre touche.
        XSLFTextBox box = slide.createTextBox();
        box.setText("Puce a gauche");
        XSLFTextParagraph para = box.getTextParagraphs().get(0);
        para.setTextAlign(TextAlign.LEFT);
        para.setLeftMargin(22.5);
        para.setIndent(-22.5);

        int fixed = RightAlignedIndentFixer.fixInheritedIndent(slide);

        assertEquals(0, fixed, "marL/indent est le mecanisme normal d'une liste a puces alignee a gauche");
        assertEquals(22.5, para.getLeftMargin());
        assertEquals(-22.5, para.getIndent());
    }

    @Test
    void fixInheritedIndent_leavesRightAlignedParagraphsWithoutIndentUntouched() {
        XSLFTextBox box = slide.createTextBox();
        box.setText("P.5");
        XSLFTextParagraph para = box.getTextParagraphs().get(0);
        para.setTextAlign(TextAlign.RIGHT);
        para.setLeftMargin(0.0);
        para.setIndent(0.0);

        int fixed = RightAlignedIndentFixer.fixInheritedIndent(slide);

        assertEquals(0, fixed, "rien a corriger si marL/indent sont deja nuls");
    }
}
