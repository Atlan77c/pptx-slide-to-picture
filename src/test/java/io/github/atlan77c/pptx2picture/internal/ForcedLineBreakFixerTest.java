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
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * Teste {@link ForcedLineBreakFixer} independamment de toute mesure de
 * largeur reelle (qui necessite {@code Graphics2D}/{@code DrawFactory}, voir
 * {@code org.apache.poi.sl.draw.HorizontalOverflowLineBreakFixer}, teste a
 * part) : {@link ForcedLineBreakFixer#locateSplitPoint}/{@code applySplit}
 * ne dependent que du texte concatene des runs, jamais d'un rendu.
 *
 * <p>Note d'execution : comme les autres tests de ce projet dependant d'Apache
 * POI (voir {@link FontSubstitutionFixerTest}), non executes cote agent
 * (Maven Central inaccessible dans ce bac a sable) - compile-verifie
 * manuellement contre le jar POI 5.2.5 reel avant commit.
 */
class ForcedLineBreakFixerTest {

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
    void locateSplitPoint_findsOffset_beforeLastWord() {
        XSLFTextParagraph para = addParagraph("le chat mange la souris");

        ForcedLineBreakFixer.SplitPoint sp = ForcedLineBreakFixer.locateSplitPoint(para, "le chat mange la souris", 1);

        assertTrue(sp != null);
        assertEquals(0, sp.splitRunIdx);
        // L'espace precedant le mot retire est exclu de la partie conservee (voir
        // Javadoc de findSplitOffset) : offset juste apres "la", sans l'espace suivant.
        assertEquals("le chat mange la".length(), sp.localOffsetInRun);
    }

    @Test
    void locateSplitPoint_returnsNull_whenNotEnoughWords() {
        XSLFTextParagraph para = addParagraph("un mot");

        ForcedLineBreakFixer.SplitPoint sp = ForcedLineBreakFixer.locateSplitPoint(para, "un mot", 5);

        assertNull(sp);
    }

    @Test
    void locateSplitPoint_returnsNull_whenLineTextNotFoundInParagraph() {
        XSLFTextParagraph para = addParagraph("le chat mange la souris");

        ForcedLineBreakFixer.SplitPoint sp = ForcedLineBreakFixer.locateSplitPoint(para, "texte absent du paragraphe", 1);

        assertNull(sp);
    }

    @Test
    void applySplit_movesLastWord_toNewRunAfterForcedBreak_preservingFormatting() {
        XSLFTextParagraph para = addParagraph("le chat mange la souris");
        XSLFTextRun originalRun = para.getTextRuns().get(0);
        originalRun.setFontSize(18.0);
        originalRun.setBold(true);

        ForcedLineBreakFixer.SplitPoint sp = ForcedLineBreakFixer.locateSplitPoint(para, "le chat mange la souris", 1);
        boolean applied = ForcedLineBreakFixer.applySplit(para, sp);

        assertTrue(applied);
        List<XSLFTextRun> runs = para.getTextRuns();
        // Run tronque (sans l'espace de separation) + saut de ligne + nouveau run
        // reprenant tel quel le suffixe du run d'origine (l'espace de separation
        // "voyage" avec le mot pousse a la ligne suivante - sans consequence
        // visuelle, un espace en tout debut de ligne n'etant jamais visible).
        assertEquals("le chat mange la", runs.get(0).getRawText());
        assertEquals(" souris", runs.get(runs.size() - 1).getRawText());
        assertEquals(18.0, runs.get(runs.size() - 1).getFontSize());
        assertTrue(runs.get(runs.size() - 1).isBold());
    }

    @Test
    void locateSplitPoint_returnsNull_whenDroppingWouldLeaveNothingBeforeSplit() {
        // Cas limite reel : la ligne en depassement ne contient QU'UN SEUL mot (ou
        // wordsToDrop couvre la totalite des mots de la ligne). Le point de coupure
        // calcule tombe alors a l'offset 0 (rien a conserver avant le saut de ligne),
        // et locateSplitPoint le rejette explicitement (localSplit <= 0) plutot que
        // de produire un saut de ligne degenere (ligne d'origine totalement vide) -
        // limite connue et deja documentee (voir bug_debordement_horizontal...md,
        // point "gerer le cas ou meme le mot le plus court d'une ligne ne suffit pas").
        XSLFTextParagraph para = addParagraph("mot");

        ForcedLineBreakFixer.SplitPoint sp = ForcedLineBreakFixer.locateSplitPoint(para, "mot", 1);

        assertNull(sp);
    }

    @Test
    void insertBreakBeforeLastWords_handlesMultipleSuccessiveBreaks_onSameParagraph() {
        // Reproduit le scenario du bug "mot orphelin" (slide 68, Rectangle 10) :
        // plusieurs coupures successives sur le meme paragraphe doivent
        // correctement REJOUER le(s) saut(s) de ligne deja insere(s), jamais les
        // traiter comme un run de texte ordinaire (XSLFLineBreak#setText() leve
        // toujours IllegalStateException).
        XSLFTextParagraph para = addParagraph("un deux trois quatre cinq six");

        boolean first = ForcedLineBreakFixer.insertBreakBeforeLastWords(para, "un deux trois quatre cinq six", 2);
        assertTrue(first);
        // Deuxieme coupure sur ce qui reste desormais AVANT le premier saut de ligne
        // force (simule une redetection de depassement sur une ligne differente).
        boolean second = ForcedLineBreakFixer.insertBreakBeforeLastWords(para, "un deux trois quatre", 1);
        assertTrue(second);

        // Reconstitue le texte final (les sauts de ligne ne portent pas de texte propre).
        StringBuilder full = new StringBuilder();
        for (XSLFTextRun r : para.getTextRuns()) {
            String t = r.getRawText();
            if (t != null) {
                full.append(t).append('|');
            }
        }
        String result = full.toString();
        assertTrue(result.contains("cinq six"), "le contenu du premier saut de ligne doit avoir ete prealablement rejoue intact : " + result);
        assertTrue(result.contains("quatre"), "le contenu du second saut de ligne doit etre present : " + result);
    }

    private XSLFTextParagraph addParagraph(String text) {
        XSLFTextBox box = slide.createTextBox();
        box.setText(text);
        return box.getTextParagraphs().get(0);
    }
}
