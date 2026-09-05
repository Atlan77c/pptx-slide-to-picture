package io.github.atlan77c.pptx2picture.internal;

import org.apache.poi.sl.usermodel.ShapeType;
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
import org.openxmlformats.schemas.presentationml.x2006.main.CTShape;

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
    void isCalloutShape_trueForWedgeRoundRectCallout() {
        // "Bulle narrative : rectangle a coins arrondis" - nom par defaut donne par
        // PowerPoint (FR) a ce type de forme, motif reel du slide 4 (voir Javadoc de
        // la classe, section "Exemption des bulles narratives (callouts)...").
        XSLFTextBox callout = textBox(0, 0, 100, 50, "~4 ou 35 ? dossiers par departement");
        callout.setShapeType(ShapeType.WEDGE_ROUND_RECT_CALLOUT);

        assertTrue(OverflowAwareTextFitter.isCalloutShape(callout));
    }

    @Test
    void isCalloutShape_trueForOtherCalloutFamilyMembers() {
        XSLFTextBox cloudCallout = textBox(0, 0, 100, 50, "Autre variante de bulle");
        cloudCallout.setShapeType(ShapeType.CLOUD_CALLOUT);

        assertTrue(OverflowAwareTextFitter.isCalloutShape(cloudCallout),
                "toute la famille 'callout' de ShapeType doit etre couverte, pas seulement WEDGE_ROUND_RECT_CALLOUT");
    }

    @Test
    void isCalloutShape_falseForOrdinaryTextBox() {
        XSLFTextBox ordinary = textBox(0, 100, 200, 10, "Zone de texte ordinaire");

        assertFalse(OverflowAwareTextFitter.isCalloutShape(ordinary));
    }

    @Test
    void overflowCollidesWithText_ignoresCalloutNeighbour() {
        // Voir Javadoc de la classe, section "Exemption des bulles narratives
        // (callouts) de tout chevauchement" : l'ancre d'une bulle de rappel englobe
        // geometriquement son bec, jusqu'a l'endroit qu'il designe - un "chevauchement"
        // purement geometrique avec elle ne doit donc plus jamais etre remonte comme
        // une collision, meme quand l'ancre intersecte reellement la zone testee.
        XSLFTextBox self = textBox(0, 100, 200, 10, "Ce texte va deborder largement de sa boite d'origine");
        XSLFTextBox callout = textBox(0, 110, 200, 50, "~4 ou 35 ? dossiers par departement");
        callout.setShapeType(ShapeType.WEDGE_ROUND_RECT_CALLOUT);
        List<XSLFShape> all = List.of(self, callout);

        List<Rectangle2D> overflow = List.of(new Rectangle2D.Double(0, 110, 200, 15));

        assertFalse(OverflowAwareTextFitter.overflowCollidesWithText(overflow, self, all),
                "une bulle narrative (callout) ne doit jamais etre retenue comme forme voisine 'colliding'");
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

    // --- Elargissement cible aux diapositives "sommaire"/"table des matieres" (2026-08-31) ---
    // Voir Javadoc de la classe, section "Elargissement cible...".

    @Test
    void fitOverflowingText_sommaireSlideTitle_doesNotForceUnconditionalShrink_whenNoRealCollision() throws IOException {
        // Reproduit le bug racine du slide 2 (sommaire) du fichier "Refonte BEL" :
        // getTextAutofit() de POI ne suit jamais l'heritage de placeholder et
        // retombe sur NORMAL (retrecissement systematique, sans verification de
        // collision) des qu'aucun autofit n'est declare localement au niveau slide -
        // motif OOXML pourtant tres courant (et le cas de tout XSLFTextBox
        // fraichement cree). Sur une diapositive de sommaire/table des matieres,
        // ce cas n'est desormais plus jamais retreci (voir Javadoc de la classe,
        // traitement revu le 2026-08-31).
        //
        // Meme geometrie (20 lignes a 14pt dans une boite de 20pt de haut) que
        // fitOverflowingText_restoresOriginalSize_whenForcedShrinkNeverConverges
        // plus haut dans ce fichier - garantit un debordement reel quelle que
        // soit la police resolue sur la machine d'execution, SANS jamais
        // declencher le garde-fou "taille de police declaree > hauteur de la
        // boite" (14 < 20 ; contrairement a une boite d'une seule ligne trop
        // etroite, qui declencherait ce garde-fou distinct - non teste ici).
        // Aucune autre forme a proximite : precondition "pas de collision".
        XSLFTextBox title = textBox(0, -200, 300, 100, "Sommaire");
        setShapeName(title, "Titre 1");

        XSLFTextBox item = slide.createTextBox();
        item.setAnchor(new Rectangle2D.Double(10, 20, 300, 20));
        item.getTextParagraphs().get(0).getTextRuns().get(0).setText("Ligne 0");
        item.getTextParagraphs().get(0).getTextRuns().get(0).setFontSize(14.0);
        for (int i = 1; i < 20; i++) {
            XSLFTextParagraph para = item.addNewTextParagraph();
            XSLFTextRun run = para.addNewTextRun();
            run.setText("Ligne " + i);
            run.setFontSize(14.0);
        }

        BufferedImage img = new BufferedImage(10, 10, BufferedImage.TYPE_INT_ARGB);
        Graphics2D graphics = img.createGraphics();
        try {
            Rectangle2D originalAnchor = item.getAnchor();
            double measuredTextHeight = item.getTextHeight(graphics);
            assertTrue(measuredTextHeight > originalAnchor.getHeight(),
                    "precondition du test : le texte doit deborder de l'anchor d'origine");

            int changed = OverflowAwareTextFitter.fitOverflowingText(slide, graphics);

            assertEquals(0, changed, "sans collision reelle avec une autre forme, l'item ne doit pas etre "
                    + "retreci meme si son autofit n'est pas declare localement");
            for (XSLFTextParagraph para : item.getTextParagraphs()) {
                for (XSLFTextRun run : para.getTextRuns()) {
                    assertEquals(14.0, run.getFontSize(), 0.001, "la taille de police ne doit pas avoir change");
                }
            }
        } finally {
            graphics.dispose();
        }
    }

    @Test
    void fitOverflowingText_sommaireSlideTitle_neverShrinks_evenWhenRealCollisionExists() throws IOException {
        // Complement du test precedent, traitement revu le 2026-08-31 (voir Javadoc
        // de la classe) : l'elargissement n'est PLUS "collision-gated" - une forme
        // dont l'autofit n'est pas declare localement, sur une slide de sommaire,
        // n'est jamais retrecie, MEME lorsqu'une collision reelle avec une autre
        // forme de texte est detectee (contrairement a une forme noAutofit
        // classique, forcedByDeclaredNone, dont le comportement collision-gated
        // reste lui inchange - voir les tests plus haut dans ce fichier). Motif
        // reel ayant motive ce changement : au sein d'une meme liste a police
        // heritee identique, faire dependre le retrecissement d'une collision
        // propre a chaque forme produisait une incoherence visuelle entre items
        // pourtant censes s'afficher de facon uniforme (voir Javadoc de la classe).
        //
        // Meme geometrie et meme technique de positionnement dynamique de la forme
        // voisine que
        // fitOverflowingText_shrinksOnlyUntilCollisionClears_whenFullBoxFitIsUnreachable
        // plus haut dans ce fichier (calibre a partir de mesures reelles plutot que
        // de coordonnees fixes, pour rester independant de la police effectivement
        // resolue sur la machine d'execution) - garantit une collision reelle et
        // mesurable, precisement le cas que ce test verifie comme n'etant plus
        // suffisant pour declencher un retrecissement.
        XSLFTextBox title = textBox(0, -200, 300, 100, "Sommaire");
        setShapeName(title, "Titre 1");

        XSLFTextBox item = slide.createTextBox();
        item.setAnchor(new Rectangle2D.Double(0, 100, 200, 20));
        item.getTextParagraphs().get(0).getTextRuns().get(0).setText("Ligne 0");
        item.getTextParagraphs().get(0).getTextRuns().get(0).setFontSize(14.0);
        for (int i = 1; i < 20; i++) {
            XSLFTextParagraph para = item.addNewTextParagraph();
            XSLFTextRun run = para.addNewTextRun();
            run.setText("Ligne " + i);
            run.setFontSize(14.0);
        }

        BufferedImage img = new BufferedImage(10, 10, BufferedImage.TYPE_INT_ARGB);
        Graphics2D graphics = img.createGraphics();
        try {
            Rectangle2D anchor = item.getAnchor();
            double fullHeight = item.getTextHeight(graphics);
            List<Rectangle2D> fullZone = OverflowAwareTextFitter.computeOverflowZones(
                    anchor, fullHeight, item.getVerticalAlignment());
            assertFalse(fullZone.isEmpty(), "precondition : 20 lignes a 14pt doivent deborder d'une boite de 20pt");

            double neighbourTop = anchor.getY() + anchor.getHeight() + 3;
            XSLFTextBox neighbour = textBox(0, neighbourTop, 200, 100, "Forme voisine avec du texte visible");
            assertTrue(fullZone.get(0).intersects(neighbour.getAnchor()),
                    "precondition : a taille pleine, la zone de debordement doit chevaucher la forme voisine");

            int changed = OverflowAwareTextFitter.fitOverflowingText(slide, graphics);

            assertEquals(0, changed, "sur une slide de sommaire, une collision reelle ne doit plus declencher "
                    + "de retrecissement (voir Javadoc de la classe, traitement revu le 2026-08-31)");
            double sizeAfter = item.getTextParagraphs().get(0).getTextRuns().get(0).getFontSize();
            assertEquals(14.0, sizeAfter, 0.001, "la taille de police ne doit pas avoir change");
        } finally {
            graphics.dispose();
        }
    }

    @Test
    void fitOverflowingText_nonSommaireSlideTitle_stillShrinksUnconditionally_evenWithoutCollision() throws IOException {
        // Verifie la portee volontairement etroite de l'elargissement (voir
        // Javadoc de la classe) : sur une slide dont le titre NE correspond PAS a
        // un intitule de sommaire/table des matieres connu, le comportement
        // preexistant (retrecissement systematique des formes a autofit non
        // declare, sans verification de collision) doit rester inchange.
        XSLFTextBox title = textBox(0, -200, 300, 100, "Introduction");
        setShapeName(title, "Titre 1");

        XSLFTextBox item = textBox(10, 20, 300, 5, "Texte isole qui deborde de sa boite, sans aucune forme voisine");
        item.getTextParagraphs().get(0).getTextRuns().get(0).setFontSize(14.0);

        BufferedImage img = new BufferedImage(10, 10, BufferedImage.TYPE_INT_ARGB);
        Graphics2D graphics = img.createGraphics();
        try {
            Rectangle2D originalAnchor = item.getAnchor();
            double measuredTextHeight = item.getTextHeight(graphics);
            assertTrue(measuredTextHeight > originalAnchor.getHeight(),
                    "precondition du test : le texte doit deborder de l'anchor d'origine");

            int changed = OverflowAwareTextFitter.fitOverflowingText(slide, graphics);

            assertEquals(1, changed, "hors d'une slide de sommaire/table des matieres, le retrecissement "
                    + "systematique (comportement preexistant) doit rester applique meme sans collision");
            assertTrue(item.getTextParagraphs().get(0).getTextRuns().get(0).getFontSize() < 14.0,
                    "la police doit avoir ete retrecie");
        } finally {
            graphics.dispose();
        }
    }

    // --- Troisieme garde-fou (NONE) : paragraphe vide, amendement du 2026-09-01 ---
    // Voir Javadoc de la classe, section "Amendement (2026-09-01)".

    @Test
    void fitOverflowingText_declaredNone_preservesMidBlockBlankParagraphContribution() throws IOException {
        // Reproduit la regression du slide 4 d'un fichier de test reel ("fichier-test-A.pptx") (forme
        // "ZoneTexte 23") : un paragraphe vide EN MILIEU de bloc (au moins un
        // paragraphe visible le suit) occupe un espace reel qui repousse le texte
        // visible qui le suit dans le rendu reel de POI. Avant l'amendement du
        // 2026-09-01, VisibleTextMeasurer l'ignorait entierement (comme tout
        // paragraphe vide), sous-estimant la hauteur reellement necessaire et
        // pouvant donc manquer une collision reelle. Structure choisie ici -
        // paragraphe vide a la fois EN MILIEU *et* EN FIN de bloc - pour exercer
        // directement VisibleTextMeasurer, qui n'est construit que lorsque le
        // DERNIER paragraphe est vide (voir hasTrailingBlankParagraph) : seul ce
        // dernier paragraphe doit desormais etre ignore de la mesure, celui du
        // milieu doit rester compte avec sa hauteur reelle.
        // Hauteur d'ancre volontairement > 14pt (contrairement a certains tests plus haut
        // dans ce fichier qui utilisent deliberement 10pt/40pt pour declencher un AUTRE
        // garde-fou, "taille de police declaree > hauteur de la boite" - non pertinent ici,
        // voir maxDeclaredFontSize dans fitOverflowingText) : sans cette marge, ce garde-fou
        // se declencherait avant meme la detection de collision que ce test veut isoler.
        XSLFTextBox item = slide.createTextBox();
        item.setAnchor(new Rectangle2D.Double(0, 100, 200, 20));
        item.setTextAutofit(TextShape.TextAutofit.NONE);
        item.getTextParagraphs().get(0).getTextRuns().get(0).setText("Ligne 0");
        item.getTextParagraphs().get(0).getTextRuns().get(0).setFontSize(14.0);
        item.addNewTextParagraph(); // paragraphe vide EN MILIEU de bloc
        XSLFTextParagraph visible2 = item.addNewTextParagraph();
        XSLFTextRun run2 = visible2.addNewTextRun();
        run2.setText("Ligne 2");
        run2.setFontSize(14.0);
        item.addNewTextParagraph(); // paragraphe vide EN FIN de bloc (declenche VisibleTextMeasurer)

        BufferedImage img = new BufferedImage(10, 10, BufferedImage.TYPE_INT_ARGB);
        Graphics2D graphics = img.createGraphics();
        try {
            Rectangle2D anchor = item.getAnchor();

            // Deux copies de calibration jetables (retirees du slide immediatement apres
            // mesure, avant l'appel reel a fitOverflowingText) pour obtenir, sans dependre
            // de la police effectivement resolue sur la machine d'execution : la hauteur
            // CORRECTE (paragraphe du milieu conserve, seul le paragraphe final exclu) et
            // la hauteur BUGUEE (les deux paragraphes vides exclus, comportement d'avant
            // l'amendement du 2026-09-01).
            XSLFTextBox correctCalibration = slide.createTextBox();
            correctCalibration.setAnchor(new Rectangle2D.Double(0, 0, 200, 10));
            correctCalibration.getTextParagraphs().get(0).getTextRuns().get(0).setText("Ligne 0");
            correctCalibration.getTextParagraphs().get(0).getTextRuns().get(0).setFontSize(14.0);
            correctCalibration.addNewTextParagraph(); // paragraphe vide du milieu, conserve ici
            XSLFTextParagraph correctVisible2 = correctCalibration.addNewTextParagraph();
            XSLFTextRun correctRun2 = correctVisible2.addNewTextRun();
            correctRun2.setText("Ligne 2");
            correctRun2.setFontSize(14.0);
            double correctHeight = correctCalibration.getTextHeight(graphics);
            slide.removeShape(correctCalibration);

            XSLFTextBox buggyCalibration = slide.createTextBox();
            buggyCalibration.setAnchor(new Rectangle2D.Double(0, 0, 200, 10));
            buggyCalibration.getTextParagraphs().get(0).getTextRuns().get(0).setText("Ligne 0");
            buggyCalibration.getTextParagraphs().get(0).getTextRuns().get(0).setFontSize(14.0);
            XSLFTextParagraph buggyVisible2 = buggyCalibration.addNewTextParagraph();
            XSLFTextRun buggyRun2 = buggyVisible2.addNewTextRun();
            buggyRun2.setText("Ligne 2");
            buggyRun2.setFontSize(14.0);
            double buggyHeight = buggyCalibration.getTextHeight(graphics);
            slide.removeShape(buggyCalibration);

            assertTrue(correctHeight > buggyHeight,
                    "precondition : conserver le paragraphe vide du milieu doit mesurer une hauteur plus grande "
                            + "qu'en l'ignorant (sa hauteur reelle est positive)");

            List<Rectangle2D> correctZone = OverflowAwareTextFitter.computeOverflowZones(
                    anchor, correctHeight, item.getVerticalAlignment());
            assertFalse(correctZone.isEmpty(), "precondition : la hauteur correcte doit deborder de l'anchor (20pt)");

            // Place le voisin exactement entre la fin de la zone buguee (plus courte) et la
            // fin de la zone correcte (plus longue) - un alignement TOP fait toujours partir
            // les deux zones du meme point (le bas de l'anchor), seule leur etendue differe -
            // ce qui garantit que la zone buguee NE chevauche PAS le voisin tandis que la
            // zone correcte le chevauche.
            List<Rectangle2D> buggyZone = OverflowAwareTextFitter.computeOverflowZones(
                    anchor, buggyHeight, item.getVerticalAlignment());
            double buggyZoneHeight = buggyZone.isEmpty() ? 0 : buggyZone.get(0).getHeight();
            double neighbourTop = anchor.getY() + anchor.getHeight()
                    + (buggyZoneHeight + correctZone.get(0).getHeight()) / 2.0;
            XSLFTextBox neighbour = textBox(0, neighbourTop, 200, 100, "Forme voisine avec du texte visible");

            assertFalse(!buggyZone.isEmpty() && buggyZone.get(0).intersects(neighbour.getAnchor()),
                    "precondition : ignorer le paragraphe vide du milieu (ancien comportement bugue) ne doit PAS "
                            + "detecter de collision avec le voisin");
            assertTrue(correctZone.get(0).intersects(neighbour.getAnchor()),
                    "precondition : conserver le paragraphe vide du milieu doit detecter une collision reelle "
                            + "avec le voisin");

            int changed = OverflowAwareTextFitter.fitOverflowingText(slide, graphics);

            assertEquals(1, changed, "la collision reelle (mesuree en conservant la contribution du paragraphe "
                    + "vide du milieu) doit etre detectee et corrigee - l'ignorer aurait manque la collision "
                    + "(regression 'fichier-test-A.pptx', voir Javadoc de la classe)");
            double sizeAfter = item.getTextParagraphs().get(0).getTextRuns().get(0).getFontSize();
            assertTrue(sizeAfter < 14.0, "la police doit avoir ete retrecie");
        } finally {
            graphics.dispose();
        }
    }

    @Test
    void fitOverflowingText_declaredNone_stillIgnoresTrailingOnlyBlankParagraph() throws IOException {
        // Non-regression du cas d'origine ("Troisieme garde-fou", item2 du sommaire
        // "Refonte BEL") : un paragraphe vide UNIQUEMENT en fin de bloc (aucun
        // paragraphe vide en milieu de bloc) doit continuer a etre entierement
        // ignore de la mesure - l'amendement du 2026-09-01 (voir Javadoc de la
        // classe et le test precedent) ne doit rien changer a ce cas precis. Comme
        // pour ce test precedent, calibre un voisin positionne de sorte que SEULE
        // la hauteur incluant (a tort) le paragraphe vide final chevaucherait le
        // voisin - la hauteur correcte (paragraphe final ignore) ne doit pas le
        // chevaucher - pour verifier positivement l'exclusion plutot que de se
        // contenter de l'absence de tout voisin.
        // Hauteur d'ancre volontairement > 14pt, meme raison que dans le test precedent :
        // eviter de declencher le garde-fou "taille de police declaree > hauteur de la
        // boite" (maxDeclaredFontSize), qui ferait sortir avant meme la detection de
        // collision et rendrait ce test vrai pour une mauvaise raison (changed == 0 par
        // court-circuit plutot que par exclusion correcte du paragraphe vide final).
        XSLFTextBox item = slide.createTextBox();
        item.setAnchor(new Rectangle2D.Double(0, 100, 200, 20));
        item.setTextAutofit(TextShape.TextAutofit.NONE);
        item.getTextParagraphs().get(0).getTextRuns().get(0).setText("Ligne 0");
        item.getTextParagraphs().get(0).getTextRuns().get(0).setFontSize(14.0);
        item.addNewTextParagraph(); // paragraphe vide EN FIN de bloc uniquement

        BufferedImage img = new BufferedImage(10, 10, BufferedImage.TYPE_INT_ARGB);
        Graphics2D graphics = img.createGraphics();
        try {
            Rectangle2D anchor = item.getAnchor();

            // Hauteur "correcte" (paragraphe final ignore) : mesuree sur une copie de
            // calibration jetable ne contenant que "Ligne 0", retiree du slide aussitot
            // apres mesure.
            XSLFTextBox correctCalibration = slide.createTextBox();
            correctCalibration.setAnchor(new Rectangle2D.Double(0, 0, 200, 10));
            correctCalibration.getTextParagraphs().get(0).getTextRuns().get(0).setText("Ligne 0");
            correctCalibration.getTextParagraphs().get(0).getTextRuns().get(0).setFontSize(14.0);
            double correctHeight = correctCalibration.getTextHeight(graphics);
            slide.removeShape(correctCalibration);

            // Hauteur "buguee" (paragraphe final compte a tort) : mesuree directement sur
            // item, qui contient encore son paragraphe vide final a ce stade.
            double fullHeight = item.getTextHeight(graphics);

            assertTrue(fullHeight > correctHeight,
                    "precondition : compter a tort le paragraphe vide final doit mesurer une hauteur plus grande");

            List<Rectangle2D> fullZone = OverflowAwareTextFitter.computeOverflowZones(
                    anchor, fullHeight, item.getVerticalAlignment());
            assertFalse(fullZone.isEmpty(), "precondition : la hauteur buguee doit deborder de l'anchor (20pt)");

            List<Rectangle2D> correctZone = OverflowAwareTextFitter.computeOverflowZones(
                    anchor, correctHeight, item.getVerticalAlignment());
            double correctZoneHeight = correctZone.isEmpty() ? 0 : correctZone.get(0).getHeight();

            // Voisin place entre la fin de la zone correcte (plus courte) et la fin de la
            // zone buguee (plus longue) - meme technique que le test precedent.
            double neighbourTop = anchor.getY() + anchor.getHeight()
                    + (correctZoneHeight + fullZone.get(0).getHeight()) / 2.0;
            XSLFTextBox neighbour = textBox(0, neighbourTop, 200, 100, "Forme voisine avec du texte visible");

            assertFalse(!correctZone.isEmpty() && correctZone.get(0).intersects(neighbour.getAnchor()),
                    "precondition : ignorer le paragraphe vide final (comportement attendu) ne doit PAS detecter "
                            + "de collision avec le voisin");
            assertTrue(fullZone.get(0).intersects(neighbour.getAnchor()),
                    "precondition : compter a tort le paragraphe vide final detecterait, lui, une collision "
                            + "avec le voisin");

            int changed = OverflowAwareTextFitter.fitOverflowingText(slide, graphics);

            assertEquals(0, changed, "le paragraphe vide final doit rester ignore de la mesure (comportement "
                    + "inchange par l'amendement du 2026-09-01) - la collision calculee a tort en le comptant ne "
                    + "doit pas declencher de retrecissement");
            assertEquals(14.0, item.getTextParagraphs().get(0).getTextRuns().get(0).getFontSize(), 0.001,
                    "la taille de police ne doit pas avoir change");
        } finally {
            graphics.dispose();
        }
    }

    // --- Quatrieme garde-fou : point de retour a la ligne "a la limite" (2026-09-02,
    // amende le meme jour - simulation independante remplacant la technique
    // differentielle initiale) --- Voir Javadoc de la classe, section "Quatrieme
    // garde-fou...".

    @Test
    void fitOverflowingText_declaredNone_shrinksFont_whenWrapPointIsAtAnchorEdge() throws IOException {
        // Reproduit le motif general (texte simple, non gras) du slide 5 d'un fichier
        // de test reel ("fichier-test-A.pptx", ZoneTexte 41) : calibre, par recherche dichotomique sur la
        // largeur reelle, la largeur exacte de bascule entre 1 et 2 lignes selon la
        // mesure NATIVE de POI, puis place l'ancre de la forme tout juste au-dessus de
        // ce seuil : a la largeur PLEINE, le texte tient sur 1 ligne (aucun
        // debordement vertical, donc rien qui declencherait le reste de cette methode
        // independamment de ce garde-fou) ; la marge de securite (WIDTH_SAFETY_MARGIN,
        // appliquee a l'interieur de la simulation independante) doit suffire a faire
        // basculer la mesure a largeur reduite sur 2 lignes - exactement le motif que
        // ce garde-fou doit detecter et corriger, meme sans le biais specifique au gras
        // (couvert separement ci-dessous). Famille de police fixee explicitement
        // (plutot que laissee heritee/nulle) pour eviter toute source de bruit dans la
        // mesure NATIVE de POI (utilisee ici pour le calibrage ET pour les deux appels
        // de {@code isWrapMarginUnstable}, qui mesurent tous deux sur le MEME objet
        // {@code item} - voir Javadoc de la classe, section "Quatrieme garde-fou").
        XSLFTextBox item = slide.createTextBox();
        item.setTextAutofit(TextShape.TextAutofit.NONE);
        item.setWordWrap(true);
        XSLFTextRun run = item.getTextParagraphs().get(0).getTextRuns().get(0);
        run.setText("mot0 mot1 mot2 mot3 mot4 mot5 mot6 mot7");
        run.setFontSize(14.0);
        run.setFontFamily("SansSerif");

        BufferedImage img = new BufferedImage(10, 10, BufferedImage.TYPE_INT_ARGB);
        Graphics2D graphics = img.createGraphics();
        try {
            item.setAnchor(new Rectangle2D.Double(0, 0, 2000, 200));
            double oneLineHeight = item.getTextHeight(graphics);

            double lo = 20;
            double hi = 2000;
            item.setAnchor(new Rectangle2D.Double(0, 0, lo, 200));
            assertTrue(item.getTextHeight(graphics) > oneLineHeight + 0.01,
                    "precondition : une largeur de 20pt doit forcer un retour a la ligne");

            for (int i = 0; i < 40; i++) {
                double mid = (lo + hi) / 2.0;
                item.setAnchor(new Rectangle2D.Double(0, 0, mid, 200));
                double h = item.getTextHeight(graphics);
                if (h > oneLineHeight + 0.01) {
                    lo = mid;
                } else {
                    hi = mid;
                }
            }
            // hi (a une precision extreme pres) est la largeur minimale a 1 ligne.
            double thresholdWidth = hi;
            assertTrue(thresholdWidth > 50,
                    "sanity : le seuil calibre doit etre substantiel pour que la marge de securite "
                            + "(pourcentage de la largeur) depasse le bruit de mesure");

            item.setAnchor(new Rectangle2D.Double(0, 100, thresholdWidth + 1.0, 200));

            assertTrue(OverflowAwareTextFitter.isWrapMarginUnstable(item, item.getAnchor(), graphics),
                    "precondition : a cette largeur calibree tout juste au-dessus du seuil, la marge de "
                            + "securite appliquee a la largeur reduite doit faire basculer le texte sur une ligne "
                            + "supplementaire");

            int changed = OverflowAwareTextFitter.fitOverflowingText(slide, graphics);

            assertEquals(1, changed, "le point de retour a la ligne calibre a la limite doit declencher un "
                    + "retrecissement (regression 'fichier-test-A.pptx', voir Javadoc de la classe)");
            assertTrue(run.getFontSize() < 14.0, "la police doit avoir ete retrecie");
            assertFalse(OverflowAwareTextFitter.isWrapMarginUnstable(item, item.getAnchor(), graphics),
                    "apres retrecissement, le point de retour a la ligne doit etre redevenu stable");
        } finally {
            graphics.dispose();
        }
    }

    @Test
    void fitOverflowingText_declaredNone_shrinksFont_whenBoldWrapPointIsAtAnchorEdge() throws IOException {
        // Reproduit specifiquement l'angle mort ayant motive l'amendement du
        // 2026-09-02 (voir Javadoc de la classe) : le run concerne est en GRAS, comme
        // le groupe de runs "outil -> ... solution Xpé " de "ZoneTexte 41". Meme
        // technique de calibration que le test precedent, mais la recherche
        // dichotomique porte cette fois directement sur le texte EN GRAS (la police
        // est deja en gras au moment de calibrer via item.getTextHeight()) : le seuil
        // trouve est donc le seuil 1/2 lignes REEL pour ce texte gras, mesure par POI
        // lui-meme - independamment de la marge supplementaire ajoutee par ce
        // garde-fou, qui mesure la surlargeur du mot gras via
        // Graphics2D.getFontMetrics() (police en gras vs. la meme police sans le bit
        // gras) et la soustrait de la largeur reduite testee (voir
        // maxBoldWidthPremium, Javadoc de la classe, section "Quatrieme garde-fou").
        // Valide que ce chemin specifique au gras fonctionne de bout en bout :
        // detection au seuil, retrecissement, puis stabilisation - sans necessairement
        // reproduire l'incoherence interne precise de POI (biais entre decision de
        // coupure et dessin) qui a motive la regression sur le fichier reel, verifiee
        // elle par l'utilisateur via mvn verify et comparaison visuelle avec
        // PowerPoint.
        XSLFTextBox item = slide.createTextBox();
        item.setTextAutofit(TextShape.TextAutofit.NONE);
        item.setWordWrap(true);
        XSLFTextRun run = item.getTextParagraphs().get(0).getTextRuns().get(0);
        run.setText("mot0 mot1 mot2 mot3 mot4 mot5 mot6 mot7");
        run.setFontSize(14.0);
        run.setFontFamily("SansSerif"); // voir commentaire du test precedent : evite toute divergence de resolution de police entre les deux mesures comparees
        run.setBold(true);

        BufferedImage img = new BufferedImage(10, 10, BufferedImage.TYPE_INT_ARGB);
        Graphics2D graphics = img.createGraphics();
        try {
            item.setAnchor(new Rectangle2D.Double(0, 0, 2000, 200));
            double oneLineHeight = item.getTextHeight(graphics);

            double lo = 20;
            double hi = 2000;
            item.setAnchor(new Rectangle2D.Double(0, 0, lo, 200));
            assertTrue(item.getTextHeight(graphics) > oneLineHeight + 0.01,
                    "precondition : une largeur de 20pt doit forcer un retour a la ligne");

            for (int i = 0; i < 40; i++) {
                double mid = (lo + hi) / 2.0;
                item.setAnchor(new Rectangle2D.Double(0, 0, mid, 200));
                double h = item.getTextHeight(graphics);
                if (h > oneLineHeight + 0.01) {
                    lo = mid;
                } else {
                    hi = mid;
                }
            }
            double thresholdWidth = hi;
            assertTrue(thresholdWidth > 50,
                    "sanity : le seuil calibre doit etre substantiel pour que la marge de securite "
                            + "(pourcentage de la largeur) depasse le bruit de mesure");

            item.setAnchor(new Rectangle2D.Double(0, 100, thresholdWidth + 1.0, 200));

            assertTrue(OverflowAwareTextFitter.isWrapMarginUnstable(item, item.getAnchor(), graphics),
                    "precondition : a cette largeur calibree (texte gras) tout juste au-dessus du seuil, la "
                            + "marge de securite (plus, le cas echeant, la surprime liee au gras) appliquee a la "
                            + "largeur reduite doit faire basculer le texte sur une ligne supplementaire");

            int changed = OverflowAwareTextFitter.fitOverflowingText(slide, graphics);

            assertEquals(1, changed, "le point de retour a la ligne calibre a la limite doit declencher un "
                    + "retrecissement, meme pour un run en gras");
            assertTrue(run.getFontSize() < 14.0, "la police doit avoir ete retrecie");
            assertTrue(run.isBold(), "le run doit rester en gras - seule la taille change");
            assertFalse(OverflowAwareTextFitter.isWrapMarginUnstable(item, item.getAnchor(), graphics),
                    "apres retrecissement, le point de retour a la ligne doit etre redevenu stable");
        } finally {
            graphics.dispose();
        }
    }

    @Test
    void fitOverflowingText_declaredNone_leavesFontUnchanged_whenWrapMarginIsComfortable() throws IOException {
        // Non-regression : une largeur d'ancre tres confortable (loin du seuil de
        // bascule 1/2 lignes) ne doit jamais etre retrecie par ce garde-fou - il ne
        // doit se declencher que pour un point de coupure reellement "a la limite".
        XSLFTextBox item = slide.createTextBox();
        item.setTextAutofit(TextShape.TextAutofit.NONE);
        item.setWordWrap(true);
        XSLFTextRun run = item.getTextParagraphs().get(0).getTextRuns().get(0);
        run.setText("mot0 mot1 mot2 mot3 mot4 mot5 mot6 mot7");
        run.setFontSize(14.0);
        run.setFontFamily("SansSerif");

        BufferedImage img = new BufferedImage(10, 10, BufferedImage.TYPE_INT_ARGB);
        Graphics2D graphics = img.createGraphics();
        try {
            item.setAnchor(new Rectangle2D.Double(0, 100, 2000, 200));

            assertFalse(OverflowAwareTextFitter.isWrapMarginUnstable(item, item.getAnchor(), graphics),
                    "precondition : une largeur tres genereuse ne doit jamais etre \"a la limite\"");

            int changed = OverflowAwareTextFitter.fitOverflowingText(slide, graphics);

            assertEquals(0, changed, "une marge confortable ne doit declencher aucun retrecissement");
            assertEquals(14.0, run.getFontSize(), 0.001, "la taille de police ne doit pas avoir change");
        } finally {
            graphics.dispose();
        }
    }

    // --- Cinquieme garde-fou : fidelite de l'interligne en pourcentage ---
    //
    // Ces deux tests couvrent uniquement les gardes-fous DETERMINISTES (aucune
    // valeur declaree, ou interligne absolu en points) : ils ne dependent d'aucune
    // metrique de police reelle et sont donc stables quel que soit l'environnement
    // d'execution. La correction ACTIVE (interligne en pourcentage effectivement
    // recalcule) depend, elle, de TextLayout.getLeading() pour la police
    // reellement resolue sur la machine d'execution (voir Javadoc de la classe,
    // "Cinquieme garde-fou...") - non testee ici pour eviter un test fragile
    // (valeur attendue dependante de la police par defaut du JDK/OS d'execution) ;
    // verifiee a la place sur un rendu reel (voir project memory du bug).

    @Test
    void correctPercentLineSpacingForFidelity_leavesUnchanged_whenNoLineSpacingDeclared() {
        XSLFTextBox box = textBox(0, 0, 300, 100, "Texte sans interligne declare localement");

        BufferedImage img = new BufferedImage(10, 10, BufferedImage.TYPE_INT_ARGB);
        Graphics2D graphics = img.createGraphics();
        try {
            OverflowAwareTextFitter.correctPercentLineSpacingForFidelity(box, graphics);

            assertEquals(null, box.getTextParagraphs().get(0).getLineSpacing(),
                    "aucun interligne declare localement (et aucun heritage dans une forme toute neuve) : rien a corriger");
        } finally {
            graphics.dispose();
        }
    }

    @Test
    void correctPercentLineSpacingForFidelity_leavesUnchanged_whenSpacingIsAbsolutePoints() {
        // Convention POI : une valeur NEGATIVE de getLineSpacing()/setLineSpacing()
        // signifie un interligne ABSOLU en points (spcPts), pas un pourcentage
        // (spcPct) - hors de portee de ce garde-fou, qui ne concerne que le repli
        // +15% de POI sur l'interligne en POURCENTAGE (voir Javadoc de la classe).
        XSLFTextBox box = textBox(0, 0, 300, 100, "Texte a interligne absolu (points)");
        box.getTextParagraphs().get(0).setLineSpacing(-24.0);

        BufferedImage img = new BufferedImage(10, 10, BufferedImage.TYPE_INT_ARGB);
        Graphics2D graphics = img.createGraphics();
        try {
            OverflowAwareTextFitter.correctPercentLineSpacingForFidelity(box, graphics);

            assertEquals(-24.0, box.getTextParagraphs().get(0).getLineSpacing(), 0.001,
                    "un interligne absolu en points ne doit jamais etre modifie par ce garde-fou");
        } finally {
            graphics.dispose();
        }
    }

    /**
     * L'API XSLF publique ne propose aucun setter pour le nom d'une forme
     * ({@code getShapeName()} n'a pas de {@code setShapeName()} correspondant) -
     * on doit donc passer par le XML sous-jacent (meme technique que {@code
     * TitleRepainterTest}).
     */
    private static void setShapeName(XSLFShape shape, String name) {
        Object xmlObject = shape.getXmlObject();
        if (xmlObject instanceof CTShape) {
            ((CTShape) xmlObject).getNvSpPr().getCNvPr().setName(name);
        }
    }
}
