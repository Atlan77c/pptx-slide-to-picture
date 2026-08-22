package io.github.atlan77c.pptx2image.internal;

import org.apache.poi.sl.usermodel.VerticalAlignment;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextBox;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.awt.geom.Rectangle2D;
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
}
