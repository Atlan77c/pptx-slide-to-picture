package io.github.atlan77c.pptx2picture.internal;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextBox;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.awt.GraphicsEnvironment;
import java.io.IOException;
import java.util.Locale;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;
import static org.junit.jupiter.api.Assumptions.assumeFalse;
import static org.junit.jupiter.api.Assumptions.assumeTrue;

/**
 * Le comportement de {@link FontSubstitutionFixer} depend, par construction,
 * des polices reellement installees sur la machine executant les tests
 * (Windows de developpement avec les polices Microsoft, serveur Linux de CI
 * sans elles...). Chaque test qui porte sur une police Microsoft precise
 * (Arial, Calibri...) utilise donc {@link org.junit.jupiter.api.Assumptions}
 * pour ne s'executer que dans l'etat d'environnement qu'il verifie
 * reellement, plutot que de supposer a tort une machine sans ces polices -
 * voir la Javadoc de la classe testee, section "Constat sur le fichier reel".
 */
class FontSubstitutionFixerTest {

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
    void fixMissingFonts_leavesRunUntouched_whenDeclaredFontIsUnknownAndHasNoSubstitute() {
        // Police totalement fictive, absente de toute machine et absente de la table de
        // correspondance : ne doit jamais etre substituee, quel que soit l'environnement.
        XSLFTextRun run = addRun("Fictive Font Name Xyz Never Installed 123");

        int fixed = FontSubstitutionFixer.fixMissingFonts(slide);

        assertEquals(0, fixed);
        assertEquals("Fictive Font Name Xyz Never Installed 123", run.getFontFamily());
    }

    @Test
    void fixMissingFonts_leavesRunUntouched_whenDeclaredFontIsAlreadyAvailable() {
        String[] available = GraphicsEnvironment.getLocalGraphicsEnvironment().getAvailableFontFamilyNames();
        assumeTrue(available.length > 0, "aucune police disponible dans cet environnement de test");
        String realFont = available[0];

        XSLFTextRun run = addRun(realFont);

        int fixed = FontSubstitutionFixer.fixMissingFonts(slide);

        assertEquals(0, fixed, "une police deja installee ne doit jamais etre remplacee");
        assertEquals(realFont, run.getFontFamily());
    }

    @Test
    void fixMissingFonts_substitutesArial_whenArialAbsentAndLiberationSansAvailable() {
        assumeFalse(isAvailable("Arial"), "Arial est installee sur cette machine : rien a substituer, cf. autre test");
        assumeTrue(isAvailable("Liberation Sans"), "substitut Liberation Sans absent de cette machine");

        XSLFTextRun run = addRun("Arial");

        int fixed = FontSubstitutionFixer.fixMissingFonts(slide);

        assertEquals(1, fixed);
        assertEquals("Liberation Sans", run.getFontFamily());
    }

    @Test
    void fixMissingFonts_leavesArialUntouched_whenArialActuallyInstalled() {
        // Cas du poste de developpement Windows avec les polices Microsoft installees :
        // le filet de securite ne doit jamais se substituer a une police reellement presente.
        assumeTrue(isAvailable("Arial"), "Arial absente de cette machine : cf. autre test");

        XSLFTextRun run = addRun("Arial");

        int fixed = FontSubstitutionFixer.fixMissingFonts(slide);

        assertEquals(0, fixed);
        assertEquals("Arial", run.getFontFamily());
    }

    @Test
    void fixMissingFonts_leavesRunUntouched_whenKnownSubstituteItselfUnavailable() {
        assumeFalse(isAvailable("Times New Roman"), "Times New Roman est installee sur cette machine");
        assumeFalse(isAvailable("Liberation Serif"), "le substitut Liberation Serif est disponible sur cette machine : cf. autre test");

        XSLFTextRun run = addRun("Times New Roman");

        int fixed = FontSubstitutionFixer.fixMissingFonts(slide);

        assertEquals(0, fixed, "ni la police declaree ni son substitut connu ne sont disponibles : ne rien inventer");
        assertEquals("Times New Roman", run.getFontFamily());
    }

    @Test
    void fixMissingFonts_isIdempotent_secondCallFixesNothingMore() {
        assumeFalse(isAvailable("Arial"), "Arial est installee sur cette machine");
        assumeTrue(isAvailable("Liberation Sans"), "substitut Liberation Sans absent de cette machine");

        addRun("Arial");

        int fixedFirst = FontSubstitutionFixer.fixMissingFonts(slide);
        int fixedSecond = FontSubstitutionFixer.fixMissingFonts(slide);

        assertEquals(1, fixedFirst);
        assertEquals(0, fixedSecond, "la police substituee est desormais disponible : rien a refaire");
    }

    @Test
    void fixMissingFonts_leavesUnsetFontFamilyUntouched() {
        // Ne fixe explicitement aucune police sur le run (herite du theme/master,
        // potentiellement null cote POI) : ne doit jamais faire echouer le correctif.
        XSLFTextBox box = slide.createTextBox();
        box.setText("Texte");
        XSLFTextParagraph para = box.getTextParagraphs().get(0);
        assertTrue(para.getTextRuns().size() >= 1);

        int fixed = FontSubstitutionFixer.fixMissingFonts(slide);

        assertEquals(0, fixed);
    }

    private static boolean isAvailable(String family) {
        String lower = family.toLowerCase(Locale.ROOT);
        for (String name : GraphicsEnvironment.getLocalGraphicsEnvironment().getAvailableFontFamilyNames()) {
            if (name.toLowerCase(Locale.ROOT).equals(lower)) {
                return true;
            }
        }
        return false;
    }

    private XSLFTextRun addRun(String fontFamily) {
        XSLFTextBox box = slide.createTextBox();
        box.setText("Texte");
        XSLFTextParagraph para = box.getTextParagraphs().get(0);
        XSLFTextRun run = para.getTextRuns().get(0);
        run.setFontFamily(fontFamily);
        return run;
    }
}
