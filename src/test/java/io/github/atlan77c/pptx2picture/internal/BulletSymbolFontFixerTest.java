package io.github.atlan77c.pptx2picture.internal;

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
 * Reproduit le motif observe sur un fichier reel (slide 9 de
 * "02.CGSI 2023 04 11 - SI Fraude_DIANE-CGSI_11avril2023_v1.0.pptx") : une
 * puce Wingdings declaree avec une liste de police de repli au format CSS
 * ({@code "Wingdings,Sans-Serif"}) - non reconnue par Apache POI, qui ne
 * remappe le caractere de la puce vers son pictogramme que pour un
 * {@code typeface} correspondant exactement (insensible a la casse) a
 * "Wingdings" ou "Symbol".
 */
class BulletSymbolFontFixerTest {

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
    void fixSymbolBulletFonts_normalizesWingdingsWithFallbackList() {
        XSLFTextBox box = slide.createTextBox();
        box.setText("Population");
        XSLFTextParagraph para = box.getTextParagraphs().get(0);
        para.setBullet(true);
        para.setBulletFont("Wingdings,Sans-Serif");
        para.setBulletCharacter("q");

        int fixed = BulletSymbolFontFixer.fixSymbolBulletFonts(slide);

        assertEquals(1, fixed);
        assertEquals("Wingdings", para.getBulletFont());
    }

    @Test
    void fixSymbolBulletFonts_normalizesSymbolWithFallbackList() {
        XSLFTextBox box = slide.createTextBox();
        box.setText("Note");
        XSLFTextParagraph para = box.getTextParagraphs().get(0);
        para.setBullet(true);
        para.setBulletFont("Symbol,Sans-Serif");
        para.setBulletCharacter("·");

        int fixed = BulletSymbolFontFixer.fixSymbolBulletFonts(slide);

        assertEquals(1, fixed);
        assertEquals("Symbol", para.getBulletFont());
    }

    @Test
    void fixSymbolBulletFonts_normalizesCaseOnlyMismatch() {
        XSLFTextBox box = slide.createTextBox();
        box.setText("Perimetre");
        XSLFTextParagraph para = box.getTextParagraphs().get(0);
        para.setBullet(true);
        para.setBulletFont("wingdings");
        para.setBulletCharacter("q");

        int fixed = BulletSymbolFontFixer.fixSymbolBulletFonts(slide);

        assertEquals(1, fixed);
        assertEquals("Wingdings", para.getBulletFont());
    }

    @Test
    void fixSymbolBulletFonts_leavesAlreadyCanonicalWingdingsUntouched() {
        // Cas deja correct dans le fichier reel (puce "Perimetre") : buFont="Wingdings"
        // exact, sans liste de repli - deja reconnu par POI, rien a corriger.
        XSLFTextBox box = slide.createTextBox();
        box.setText("Perimetre");
        XSLFTextParagraph para = box.getTextParagraphs().get(0);
        para.setBullet(true);
        para.setBulletFont("Wingdings");
        para.setBulletCharacter("q");

        int fixed = BulletSymbolFontFixer.fixSymbolBulletFonts(slide);

        assertEquals(0, fixed, "buFont deja sous sa forme canonique exacte, aucune correction necessaire");
        assertEquals("Wingdings", para.getBulletFont());
    }

    @Test
    void fixSymbolBulletFonts_leavesUnknownSymbolFontUntouched() {
        // Limite assumee (voir Javadoc de la classe) : Webdings n'est pas dans la
        // petite liste codee en dur cote POI, meme parfaitement nomme - non corrige.
        XSLFTextBox box = slide.createTextBox();
        box.setText("Icone");
        XSLFTextParagraph para = box.getTextParagraphs().get(0);
        para.setBullet(true);
        para.setBulletFont("Webdings,Sans-Serif");
        para.setBulletCharacter("a");

        int fixed = BulletSymbolFontFixer.fixSymbolBulletFonts(slide);

        assertEquals(0, fixed);
        assertEquals("Webdings,Sans-Serif", para.getBulletFont());
    }

    @Test
    void fixSymbolBulletFonts_leavesParagraphsWithoutBulletCharacterUntouched() {
        // Pas d'appel a setBullet(true)/setBulletCharacter(...) : ce paragraphe de
        // zone de texte libre (pas un placeholder) ne resout aucun buChar (ni local,
        // ni herite - le "otherStyle" par defaut d'un XMLSlideShow neuf n'a pas de
        // puce), donc buFont est sans effet ici, meme mal forme.
        XSLFTextBox box = slide.createTextBox();
        box.setText("Etape 1");
        XSLFTextParagraph para = box.getTextParagraphs().get(0);
        para.setBulletFont("Wingdings,Sans-Serif");

        int fixed = BulletSymbolFontFixer.fixSymbolBulletFonts(slide);

        assertEquals(0, fixed, "pas de buChar resolu, rien a corriger");
    }
}
