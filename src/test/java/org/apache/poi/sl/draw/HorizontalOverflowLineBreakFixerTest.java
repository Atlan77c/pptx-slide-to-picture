package org.apache.poi.sl.draw;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextBox;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.awt.Graphics2D;
import java.awt.RenderingHints;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.IOException;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * Teste {@link HorizontalOverflowLineBreakFixer} de bout en bout (via un vrai
 * {@code Graphics2D}, indispensable ici puisque la detection depend de {@code
 * DrawFactory}/{@code DrawTextParagraph#breakText}, voir Javadoc de la classe
 * testee) plutot qu'au niveau de {@link
 * io.github.atlan77c.pptx2picture.internal.ForcedLineBreakFixer} (teste
 * separement, sans dependance a un rendu - voir {@code
 * ForcedLineBreakFixerTest}).
 *
 * <p>Note d'execution : comme les autres tests de ce projet dependant d'Apache
 * POI (voir {@code FontSubstitutionFixerTest}), non executes cote agent
 * (Maven Central inaccessible dans ce bac a sable) - compile-verifie ET
 * execute manuellement (harnais de test reflexif ad hoc) contre le jar POI
 * 5.2.5 reel avant commit.
 */
class HorizontalOverflowLineBreakFixerTest {

    private XMLSlideShow ppt;
    private XSLFSlide slide;
    private BufferedImage img;
    private Graphics2D graphics;

    @BeforeEach
    void setUp() {
        ppt = new XMLSlideShow();
        slide = ppt.createSlide();
        img = new BufferedImage(200, 200, BufferedImage.TYPE_INT_ARGB);
        graphics = img.createGraphics();
        graphics.setRenderingHint(RenderingHints.KEY_TEXT_ANTIALIASING, RenderingHints.VALUE_TEXT_ANTIALIAS_ON);
    }

    @AfterEach
    void tearDown() throws IOException {
        graphics.dispose();
        ppt.close();
    }

    @Test
    void fixHorizontalOverflow_insertsForcedBreak_whenLineFallsInTheSafetyMarginDangerZone() {
        // Reproduit, independamment de toute police particuliere, la "zone
        // dangereuse" a l'origine meme de ce correctif (voir Javadoc de
        // HorizontalOverflowLineBreakFixer#WIDTH_SAFETY_MARGIN) : une ligne dont la
        // largeur mesuree se situe entre 97% et 100% de la largeur utile - POI la
        // considere comme tenant tout juste (breakText() ne la coupe pas), mais le
        // rendu reel peut encore la couper (ecart Java2D/rendu reel confirme
        // empiriquement sur la slide 57 de "Cadrage de vision_ Definir_0.6.pptx").
        //
        // Plutot que de coder en dur une largeur de boite (fragile - la meme boite
        // ne produirait pas forcement le meme ratio largeur-mesuree/largeur-utile
        // avec une police differente, ex. Windows en production vs les polices de
        // substitution de ce bac a sable, voir FontSubstitutionFixerTest), la
        // largeur cible est CALCULEE a partir d'une mesure reelle prealable de "AA
        // BB" dans CET environnement - reproductible quelle que soit la police
        // effectivement utilisee au moment de l'execution du test.
        double naturalWidth = measureNaturalWidth("AA BB");
        double targetWidth = naturalWidth / 0.985; // pile entre 97% (0.97) et 100% de la marge.
        XSLFTextBox box = slide.createTextBox();
        box.setAnchor(new Rectangle2D.Double(10, 10, targetWidth, 200));
        box.setWordWrap(true);
        box.setLeftInset(0);
        box.setRightInset(0);
        box.setText("AA BB");

        int shapesFixed = HorizontalOverflowLineBreakFixer.fixHorizontalOverflow(slide, graphics);

        assertEquals(1, shapesFixed);
        List<XSLFTextRun> runs = box.getTextParagraphs().get(0).getTextRuns();
        boolean hasLineBreak = false;
        for (XSLFTextRun r : runs) {
            if ("XSLFLineBreak".equals(r.getClass().getSimpleName())) {
                hasLineBreak = true;
            }
        }
        assertTrue(hasLineBreak, "un saut de ligne force doit avoir ete insere");
        assertEquals("AA", runs.get(0).getRawText());
        assertEquals(" BB", runs.get(runs.size() - 1).getRawText());
    }

    /**
     * Mesure la largeur NON enveloppee (une seule ligne, {@code wordWrap=false},
     * ancre tres large) de {@code text} dans l'environnement d'execution reel -
     * meme mecanisme de mesure ({@code DrawFactory}/{@code DrawTextParagraph})
     * que celui utilise par la classe testee, ce qui rend le resultat coherent
     * avec elle par construction.
     */
    private double measureNaturalWidth(String text) {
        XSLFTextBox probe = slide.createTextBox();
        probe.setAnchor(new Rectangle2D.Double(10, 10, 2000, 100));
        probe.setWordWrap(false);
        probe.setLeftInset(0);
        probe.setRightInset(0);
        probe.setText(text);
        DrawFactory fact = DrawFactory.getInstance(graphics);
        DrawTextParagraph dp = fact.getDrawable(probe.getTextParagraphs().get(0));
        dp.setFirstParagraph(true);
        dp.breakText(graphics);
        return dp.lines.get(0).getWidth();
    }

    @Test
    void fixHorizontalOverflow_leavesShapeUntouched_whenTextAlreadyFits() {
        XSLFTextBox box = slide.createTextBox();
        box.setAnchor(new Rectangle2D.Double(10, 10, 400, 100));
        box.setWordWrap(true);
        box.setText("court");

        int shapesFixed = HorizontalOverflowLineBreakFixer.fixHorizontalOverflow(slide, graphics);

        assertEquals(0, shapesFixed);
        assertEquals(1, box.getTextParagraphs().get(0).getTextRuns().size());
    }

    @Test
    void fixHorizontalOverflow_excludesShape_whenInsetsDominateWidth() {
        // Meme technique de "zone dangereuse" que le test precedent, mais avec des
        // marges internes (insets) qui a elles seules consomment plus de la moitie
        // de la largeur de l'ancre - voir Javadoc de
        // HorizontalOverflowLineBreakFixer#isExcluded (faux positif confirme par
        // l'utilisateur sur un petit connecteur, slide 90). La largeur UTILE
        // (ancre moins insets) est deliberement la meme que dans le test precedent
        // (qui, lui, produit bien une correction) : seule la dominance des insets
        // doit faire la difference.
        double naturalWidth = measureNaturalWidth("AA BB");
        double effectiveWidth = naturalWidth / 0.985;
        double inset = effectiveWidth; // insets = largeur utile chacun.
        double anchorWidth = 3 * effectiveWidth; // (inset+inset)/anchorWidth = 2/3 >= seuil de 50%.
        XSLFTextBox box = slide.createTextBox();
        box.setAnchor(new Rectangle2D.Double(10, 10, anchorWidth, 200));
        box.setWordWrap(true);
        box.setLeftInset(inset);
        box.setRightInset(inset);
        box.setText("AA BB");

        int shapesFixed = HorizontalOverflowLineBreakFixer.fixHorizontalOverflow(slide, graphics);

        assertEquals(0, shapesFixed, "forme exclue (marges internes dominantes) : aucune correction ne doit etre tentee");
        List<XSLFTextRun> runs = box.getTextParagraphs().get(0).getTextRuns();
        assertEquals(1, runs.size(), "aucun saut de ligne ne doit avoir ete insere sur une forme exclue");
    }

    @Test
    void fixHorizontalOverflow_excludesVisionCibleBadge_byExactContentMatch() {
        // Meme technique de "zone dangereuse" (voir premier test), appliquee cette
        // fois au texte EXACT du badge - faux positif confirme par l'utilisateur
        // (texte gras, mecanisme non elucide) - exclusion PONCTUELLE par contenu
        // exact, jamais generalisee.
        XSLFTextBox box = dangerZoneTextBox("Vision cible de la solution EPIC 4");

        int shapesFixed = HorizontalOverflowLineBreakFixer.fixHorizontalOverflow(slide, graphics);

        assertEquals(0, shapesFixed);
        assertEquals(1, box.getTextParagraphs().get(0).getTextRuns().size());
    }

    @Test
    void fixHorizontalOverflow_doesNotExclude_similarButDifferentText() {
        // Verifie que l'exclusion du badge est bien une correspondance EXACTE
        // (trim() uniquement) et ne s'etend pas a un texte simplement proche - meme
        // "zone dangereuse", un seul mot ajoute par rapport au texte exact du badge.
        dangerZoneTextBox("Vision cible de la solution EPIC 4 bis");

        int shapesFixed = HorizontalOverflowLineBreakFixer.fixHorizontalOverflow(slide, graphics);

        assertEquals(1, shapesFixed);
    }

    /**
     * Construit une boite de texte dont la largeur utile place {@code text} (une
     * fois enveloppe sur une seule ligne) dans la "zone dangereuse" 97%-100% de
     * {@link HorizontalOverflowLineBreakFixer#WIDTH_SAFETY_MARGIN} - voir le
     * commentaire du premier test de cette classe pour le detail du calcul.
     */
    private XSLFTextBox dangerZoneTextBox(String text) {
        double naturalWidth = measureNaturalWidth(text);
        double targetWidth = naturalWidth / 0.985;
        XSLFTextBox box = slide.createTextBox();
        box.setAnchor(new Rectangle2D.Double(10, 10, targetWidth, 200));
        box.setWordWrap(true);
        box.setLeftInset(0);
        box.setRightInset(0);
        box.setText(text);
        return box;
    }
}
