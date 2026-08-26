package io.github.atlan77c.pptx2picture.internal;

import org.apache.poi.sl.usermodel.TextShape;
import org.apache.poi.sl.usermodel.VerticalAlignment;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextBox;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
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
        // pratique sur un fichier reel a provoquer de nouveaux chevauchements avec les
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
    void fitOverflowingText_neverForcesShrink_whenDeclaredFontSizeAloneExceedsAnchorHeight() throws IOException {
        // Reproduit un motif observe sur un fichier reel : un caractere decoratif
        // unique ("*") en 72pt dans une boite noAutofit de ~40pt de haut,
        // chevauchant deliberement une forme voisine toute proche - exactement
        // comme PowerPoint l'affiche (l'auteur a choisi une police enorme dans
        // une petite boite pour dessiner un accent visuel). Meme si le
        // chevauchement geometrique est bien reel (verifie ci-dessous par la
        // precondition), aucun retrecissement ne doit etre applique : la
        // taille de police declaree (72pt) depasse deja a elle seule la
        // hauteur de la boite (40pt), donc ce n'est jamais un artefact de
        // mesure Java2D - PowerPoint montrerait le meme debordement.
        XSLFTextBox star = textBox(0, 100, 40, 40, "*");
        star.setTextAutofit(TextShape.TextAutofit.NONE);
        star.getTextParagraphs().get(0).getTextRuns().get(0).setFontSize(72.0);

        XSLFTextBox neighbour = textBox(0, 130, 200, 50, "Forme voisine avec du texte visible");

        BufferedImage img = new BufferedImage(10, 10, BufferedImage.TYPE_INT_ARGB);
        Graphics2D graphics = img.createGraphics();
        try {
            Rectangle2D starAnchor = star.getAnchor();
            Rectangle2D neighbourAnchor = neighbour.getAnchor();
            double measuredTextHeight = star.getTextHeight(graphics);
            List<Rectangle2D> overflow = OverflowAwareTextFitter.computeOverflowZones(
                    starAnchor, measuredTextHeight, star.getVerticalAlignment());
            assertFalse(overflow.isEmpty(), "precondition : le caractere a 72pt doit deborder de sa boite de 40pt");
            assertTrue(overflow.get(0).intersects(neighbourAnchor),
                    "precondition : la zone de debordement doit chevaucher reellement la forme voisine");

            int changed = OverflowAwareTextFitter.fitOverflowingText(slide, graphics);

            assertEquals(0, changed, "aucune forme ne doit etre retrecie malgre la collision reelle");
            Double fontSizeAfter = star.getTextParagraphs().get(0).getTextRuns().get(0).getFontSize();
            assertEquals(72.0, fontSizeAfter, 0.001, "la taille de police du caractere decoratif ne doit pas changer");
        } finally {
            graphics.dispose();
        }
    }

    @Test
    void fitOverflowingText_appliesSafetyMarginBelowAnchorHeight_notJustStrictlyUnder() throws IOException {
        // Reproduit un cas reel diagnostique : une forme spAutoFit calibree tres
        // exactement par PowerPoint (aucune marge native) pour laquelle un
        // retrecissement qui viserait anchor.getHeight() au plus juste laisse une
        // marge residuelle de moins d'1pt - insuffisante face a l'epaisseur du
        // trait de bordure de la forme. Le correctif retenu vise desormais une
        // hauteur legerement inferieure a l'anchor (marge de securite), verifiee
        // ici geometriquement plutot que visuellement.
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

    @Test
    void fitOverflowingText_restoresOriginalSize_whenForcedShrinkNeverConverges() throws IOException {
        // Reproduit un motif observe sur un fichier reel : une boite noAutofit
        // bien plus petite que son contenu (ici 20 courtes lignes a 14pt dans une
        // boite de seulement 20pt de haut), avec une forme voisine assez proche
        // pour que la zone de debordement mesuree la chevauche -> collision
        // "reelle" detectee, retrecissement force declenche. Contrairement au cas
        // couvert par le test appliesSafetyMarginBelowAnchorHeight ci-dessus (qui
        // converge proprement), aucune reduction de police plausible ne peut faire
        // tenir 20 lignes dans 20pt : meme ecrasee jusqu'a MIN_SCALE (25%) ou
        // MAX_ITER iterations, le texte deborde toujours trop largement. C'est le
        // signal retenu pour distinguer un veritable artefact de mesure
        // (Java2D/PowerPoint) d'une boite structurellement trop petite pour son
        // contenu (legende/annotation flottante) : dans ce second cas, la taille
        // d'origine doit etre restauree plutot que de produire un texte ecrase.
        XSLFTextBox box = slide.createTextBox();
        box.setAnchor(new Rectangle2D.Double(0, 100, 200, 20));
        box.setTextAutofit(TextShape.TextAutofit.NONE);
        box.getTextParagraphs().get(0).getTextRuns().get(0).setText("Ligne 0");
        box.getTextParagraphs().get(0).getTextRuns().get(0).setFontSize(14.0);
        for (int i = 1; i < 20; i++) {
            XSLFTextParagraph para = box.addNewTextParagraph();
            XSLFTextRun run = para.addNewTextRun();
            run.setText("Ligne " + i);
            run.setFontSize(14.0);
        }

        XSLFTextBox neighbour = textBox(0, 125, 200, 50, "Forme voisine avec du texte visible");

        BufferedImage img = new BufferedImage(10, 10, BufferedImage.TYPE_INT_ARGB);
        Graphics2D graphics = img.createGraphics();
        try {
            Rectangle2D originalAnchor = box.getAnchor();
            double measuredTextHeight = box.getTextHeight(graphics);
            List<Rectangle2D> overflow = OverflowAwareTextFitter.computeOverflowZones(
                    originalAnchor, measuredTextHeight, box.getVerticalAlignment());
            assertFalse(overflow.isEmpty(), "precondition : 20 lignes a 14pt doivent deborder d'une boite de 20pt");
            assertTrue(overflow.get(0).intersects(neighbour.getAnchor()),
                    "precondition : la zone de debordement mesuree doit chevaucher la forme voisine");

            int changed = OverflowAwareTextFitter.fitOverflowingText(slide, graphics);

            assertEquals(0, changed, "une boite qui ne peut jamais contenir son texte, meme retrecie au maximum, "
                    + "ne doit pas etre comptee comme corrigee");
            for (XSLFTextParagraph para : box.getTextParagraphs()) {
                for (XSLFTextRun run : para.getTextRuns()) {
                    assertEquals(14.0, run.getFontSize(), 0.001,
                            "la taille de police d'origine doit etre restauree, pas laissee ecrasee au minimum");
                }
            }
        } finally {
            graphics.dispose();
        }
    }

    @Test
    void fitOverflowingText_shrinksOnlyUntilCollisionClears_whenFullBoxFitIsUnreachable() throws IOException {
        // Complement du test precedent : ici, le rétrécissement complet dans la boîte
        // est tout aussi hors d'atteinte (20 lignes dans 20pt de haut), mais la forme
        // voisine est placée assez loin pour qu'un rétrécissement modéré (pas un
        // écrasement au minimum) suffise à faire disparaître la collision réelle. La
        // deuxième passe doit alors s'arrêter dès que la collision disparaît, sans
        // chercher à tout prix un ajustement complet dans la boîte (hors d'atteinte) ni
        // restaurer la taille d'origine (puisqu'une collision réelle a bien été
        // résolue) - un débordement résiduel au-delà de la boîte reste accepté.
        XSLFTextBox box = slide.createTextBox();
        box.setAnchor(new Rectangle2D.Double(0, 100, 200, 20));
        box.setTextAutofit(TextShape.TextAutofit.NONE);
        box.getTextParagraphs().get(0).getTextRuns().get(0).setText("Ligne 0");
        box.getTextParagraphs().get(0).getTextRuns().get(0).setFontSize(14.0);
        for (int i = 1; i < 20; i++) {
            XSLFTextParagraph para = box.addNewTextParagraph();
            XSLFTextRun run = para.addNewTextRun();
            run.setText("Ligne " + i);
            run.setFontSize(14.0);
        }

        BufferedImage img = new BufferedImage(10, 10, BufferedImage.TYPE_INT_ARGB);
        Graphics2D graphics = img.createGraphics();
        try {
            Rectangle2D anchor = box.getAnchor();

            // Calibre la position de la forme voisine a partir de mesures reelles plutot
            // que de coordonnees fixes, pour rester independant de la police effectivement
            // resolue sur la machine d'execution (comme le reste de ce fichier de test).
            double fullHeight = box.getTextHeight(graphics);
            List<Rectangle2D> fullZone = OverflowAwareTextFitter.computeOverflowZones(
                    anchor, fullHeight, box.getVerticalAlignment());
            assertFalse(fullZone.isEmpty(), "precondition : 20 lignes a 14pt doivent deborder d'une boite de 20pt");

            for (XSLFTextParagraph para : box.getTextParagraphs()) {
                for (XSLFTextRun run : para.getTextRuns()) {
                    run.setFontSize(14.0 * 0.7);
                }
            }
            double shrunkHeight = box.getTextHeight(graphics);
            List<Rectangle2D> shrunkZone = OverflowAwareTextFitter.computeOverflowZones(
                    anchor, shrunkHeight, box.getVerticalAlignment());
            double shrunkZoneHeight = shrunkZone.isEmpty() ? 0 : shrunkZone.get(0).getHeight();
            for (XSLFTextParagraph para : box.getTextParagraphs()) {
                for (XSLFTextRun run : para.getTextRuns()) {
                    run.setFontSize(14.0); // restaure avant l'appel reel a fitOverflowingText
                }
            }
            assertTrue(shrunkZoneHeight < fullZone.get(0).getHeight(),
                    "precondition : rétrécir la police doit reduire la zone de debordement mesuree");

            // Hauteur genereuse (100pt) pour que cette forme voisine ne deborde jamais
            // elle-meme de son propre anchor - sinon fitOverflowingText la retrecirait
            // aussi (autofit par defaut d'une XSLFTextBox fraichement creee), faussant le
            // compte de formes retrecies verifie plus bas (on veut isoler le seul effet
            // sur `box`).
            double neighbourTop = anchor.getY() + anchor.getHeight() + shrunkZoneHeight + 3;
            XSLFTextBox neighbour = textBox(0, neighbourTop, 200, 100, "Forme voisine avec du texte visible");
            assertTrue(fullZone.get(0).intersects(neighbour.getAnchor()),
                    "precondition : a taille pleine, la zone de debordement doit chevaucher la forme voisine");

            int changed = OverflowAwareTextFitter.fitOverflowingText(slide, graphics);

            assertEquals(1, changed, "une collision reelle resolue par un retrecissement modere doit etre comptee");
            double sizeAfter = box.getTextParagraphs().get(0).getTextRuns().get(0).getFontSize();
            assertTrue(sizeAfter < 14.0, "la police doit avoir ete retrecie par rapport a l'original");
            assertTrue(sizeAfter > 14.0 * MIN_SCALE_FOR_TEST,
                    "un retrecissement modere doit suffire - pas besoin d'ecraser jusqu'a la limite basse");

            double heightAfter = box.getTextHeight(graphics);
            List<Rectangle2D> zoneAfter = OverflowAwareTextFitter.computeOverflowZones(
                    anchor, heightAfter, box.getVerticalAlignment());
            assertFalse(OverflowAwareTextFitter.overflowCollidesWithText(zoneAfter, box, List.of(box, neighbour)),
                    "la collision doit avoir disparu");
        } finally {
            graphics.dispose();
        }
    }

    /** Doit rester strictement egal a {@code MIN_SCALE} dans {@link OverflowAwareTextFitter}. */
    private static final double MIN_SCALE_FOR_TEST = 0.25;
}
