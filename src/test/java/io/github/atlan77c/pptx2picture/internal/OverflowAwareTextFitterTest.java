package io.github.atlan77c.pptx2picture.internal;

import org.apache.poi.sl.usermodel.TextShape;
import org.apache.poi.sl.usermodel.VerticalAlignment;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextBox;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.awt.Graphics2D;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.IOException;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * Tests de la logique geometrique pure (zones de debordement, detection de
 * collision) - independants de toute police installee sur la machine
 * d'execution, contrairement au calcul de {@code getTextHeight()} lui-meme
 * qui depend des metriques Java2D/AWT reelles. Les formes de texte utilisees
 * ici servent uniquement a porter un anchor et un texte ; aucun rendu de
 * texte n'est effectue.
 */
class OverflowAwareTextFitterTest {

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

    private XSLFTextBox textBox(double x, double y, double w, double h, String text) {
        XSLFTextBox box = slide.createTextBox();
        box.setAnchor(new Rectangle2D.Double(x, y, w, h));
        box.setText(text);
        return box;
    }

    @Test
    void computeOverflowZones_returnsEmptyList_whenTextFitsInAnchor() {
        Rectangle2D anchor = new Rectangle2D.Double(0, 0, 100, 50);
        List<Rectangle2D> zones = OverflowAwareTextFitter.computeOverflowZones(anchor, 40, VerticalAlignment.TOP);
        assertTrue(zones.isEmpty());
    }

    @Test
    void computeOverflowZones_topAlignment_growsDownwardOnly() {
        Rectangle2D anchor = new Rectangle2D.Double(0, 100, 200, 10);
        List<Rectangle2D> zones = OverflowAwareTextFitter.computeOverflowZones(anchor, 25, VerticalAlignment.TOP);

        assertEquals(1, zones.size());
        Rectangle2D zone = zones.get(0);
        assertEquals(110, zone.getY(), 0.001, "la zone de debordement doit commencer au bas de l'anchor");
        assertEquals(15, zone.getHeight(), 0.001, "l'exces (25-10) doit etre entierement place en dessous");
    }

    @Test
    void computeOverflowZones_bottomAlignment_growsUpwardOnly() {
        Rectangle2D anchor = new Rectangle2D.Double(0, 100, 200, 10);
        List<Rectangle2D> zones = OverflowAwareTextFitter.computeOverflowZones(anchor, 25, VerticalAlignment.BOTTOM);

        assertEquals(1, zones.size());
        Rectangle2D zone = zones.get(0);
        assertEquals(85, zone.getY(), 0.001, "la zone de debordement doit se terminer au sommet de l'anchor (100-15)");
        assertEquals(15, zone.getHeight(), 0.001);
    }

    @Test
    void computeOverflowZones_middleAlignment_growsBothWaysEqually() {
        Rectangle2D anchor = new Rectangle2D.Double(0, 100, 200, 10);
        List<Rectangle2D> zones = OverflowAwareTextFitter.computeOverflowZones(anchor, 30, VerticalAlignment.MIDDLE);

        assertEquals(2, zones.size(), "l'alignement centre doit produire une zone au-dessus et une en dessous");
        double totalExcessHeight = zones.get(0).getHeight() + zones.get(1).getHeight();
        assertEquals(20, totalExcessHeight, 0.001, "l'exces total (30-10) doit etre reparti pour moitie de chaque cote");
    }

    @Test
    void overflowCollidesWithText_detectsCollisionWithNeighbourTextShape() {
        XSLFTextBox self = textBox(0, 100, 200, 10, "Ce texte va deborder largement de sa boite d'origine");
        XSLFTextBox neighbour = textBox(0, 110, 200, 50, "Forme voisine avec du texte visible");
        List<XSLFShape> all = List.of(self, neighbour);

        List<Rectangle2D> overflow = List.of(new Rectangle2D.Double(0, 110, 200, 15));

        assertTrue(OverflowAwareTextFitter.overflowCollidesWithText(overflow, self, all));
    }

    @Test
    void overflowCollidesWithText_ignoresShapesWithoutText() {
        XSLFTextBox self = textBox(0, 100, 200, 10, "Texte qui deborde");
        XSLFTextBox emptyPanel = textBox(0, 110, 200, 50, ""); // panneau de fond, sans texte
        List<XSLFShape> all = List.of(self, emptyPanel);

        List<Rectangle2D> overflow = List.of(new Rectangle2D.Double(0, 110, 200, 15));

        assertFalse(OverflowAwareTextFitter.overflowCollidesWithText(overflow, self, all),
                "une forme sans texte (fond/panneau) ne doit jamais etre consideree comme une collision");
    }

    @Test
    void overflowCollidesWithText_returnsFalseWhenIsolated() {
        XSLFTextBox self = textBox(0, 100, 200, 10, "Texte isole, rien autour");
        XSLFTextBox farAway = textBox(500, 500, 50, 50, "Loin de toute zone de debordement");
        List<XSLFShape> all = List.of(self, farAway);

        List<Rectangle2D> overflow = List.of(new Rectangle2D.Double(0, 110, 200, 15));

        assertFalse(OverflowAwareTextFitter.overflowCollidesWithText(overflow, self, all));
    }

    @Test
    void fitOverflowingText_shapeAutofit_shrinksFontInsteadOfGrowingAnchor() throws IOException {
        // Anchor volontairement minuscule (5pt) et police large (60pt) : garantit un
        // debordement massif quelle que soit la police reellement resolue sur la
        // machine d'execution (le point teste ici est la strategie de correction,
        // pas la precision des metriques - voir le commentaire de classe).
        //
        // autofit=SHAPE est traite comme NORMAL (retrecissement systematique) plutot
        // que par agrandissement de la boite : une premiere version agrandissait
        // l'anchor, plus fidele au sens strict de cet autofit, mais observee en
        // pratique sur un vrai fichier a provoquer de nouveaux chevauchements avec les
        // formes voisines situees en dessous de la boite agrandie.
        XSLFTextBox box = textBox(10, 20, 300, 5, "Texte qui necessite bien plus de place que 5pt de haut");
        box.setTextAutofit(TextShape.TextAutofit.SHAPE);
        box.getTextParagraphs().get(0).getTextRuns().get(0).setFontSize(60.0);

        BufferedImage img = new BufferedImage(10, 10, BufferedImage.TYPE_INT_ARGB);
        Graphics2D graphics = img.createGraphics();
        try {
            Rectangle2D originalAnchor = box.getAnchor();
            double measuredTextHeight = box.getTextHeight(graphics);
            assertTrue(measuredTextHeight > originalAnchor.getHeight(), "precondition du test : le texte doit deborder de l'anchor d'origine");

            int changed = OverflowAwareTextFitter.fitOverflowingText(slide, graphics);

            assertEquals(1, changed);
            Rectangle2D anchorAfter = box.getAnchor();
            assertEquals(originalAnchor.getX(), anchorAfter.getX(), 0.001, "la position X ne doit pas changer");
            assertEquals(originalAnchor.getY(), anchorAfter.getY(), 0.001, "la position Y ne doit pas changer");
            assertEquals(originalAnchor.getWidth(), anchorAfter.getWidth(), 0.001, "la largeur ne doit pas changer");
            assertEquals(originalAnchor.getHeight(), anchorAfter.getHeight(), 0.001,
                    "la hauteur de l'anchor ne doit pas changer non plus - seule la police est retrecie");

            Double fontSizeAfter = box.getTextParagraphs().get(0).getTextRuns().get(0).getFontSize();
            assertTrue(fontSizeAfter < 60.0, "la police doit avoir ete retrecie sous sa taille d'origine (60pt)");
        } finally {
            graphics.dispose();
        }
    }

    @Test
    void fitOverflowingText_appliesSafetyMarginBelowAnchorHeight_notJustStrictlyUnder() throws IOException {
        // Reproduit le cas reel diagnostique sur le fichier de test reel (voir conversion_pptx_vers_images.md,
        // section 8) : une forme spAutoFit calibree tres exactement par PowerPoint (aucune marge
        // native) pour laquelle un retrecissement qui viserait anchor.getHeight() au plus juste
        // laisse une marge residuelle de moins d'1pt - insuffisante face a l'epaisseur du trait de
        // bordure de la forme. Le correctif retenu vise desormais une hauteur legerement inferieure
        // a l'anchor (marge de securite), verifiee ici geometriquement plutot que visuellement.
        XSLFTextBox box = textBox(10, 20, 300, 50, "Texte assez long pour deborder legerement de sa boite d'origine et necessiter un retrecissement mesurable");
        box.setTextAutofit(TextShape.TextAutofit.SHAPE);
        box.getTextParagraphs().get(0).getTextRuns().get(0).setFontSize(30.0);

        BufferedImage img = new BufferedImage(10, 10, BufferedImage.TYPE_INT_ARGB);
        Graphics2D graphics = img.createGraphics();
        try {
            Rectangle2D anchor = box.getAnchor();
            double measuredTextHeight = box.getTextHeight(graphics);
            assertTrue(measuredTextHeight > anchor.getHeight(), "precondition du test : le texte doit deborder de l'anchor d'origine");

            int changed = OverflowAwareTextFitter.fitOverflowingText(slide, graphics);
            assertEquals(1, changed);

            double textHeightAfter = box.getTextHeight(graphics);
            assertTrue(textHeightAfter < anchor.getHeight(),
                    "le texte retreci doit tenir strictement sous la hauteur de l'anchor (pas seulement <=)");
        } finally {
            graphics.dispose();
        }
    }
}
