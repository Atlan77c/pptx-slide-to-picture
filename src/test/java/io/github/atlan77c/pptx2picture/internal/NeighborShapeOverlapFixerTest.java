package io.github.atlan77c.pptx2picture.internal;

import org.apache.poi.sl.usermodel.ShapeType;
import org.apache.poi.sl.usermodel.VerticalAlignment;
import org.apache.poi.sl.usermodel.TextShape;
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
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * Comme {@code OverflowAwareTextFitterTest} : la geometrie pure (zone
 * occupee, detection de collision) est testee independamment de toute police
 * installee sur la machine d'execution. Les tests d'integration qui
 * declenchent {@link NeighborShapeOverlapFixer#fixNeighborOverlaps} calibrent
 * la position de la forme voisine a partir de mesures reelles ({@code
 * getTextHeight()} prises AVANT et APRES une reduction sondee a l'avance),
 * exactement comme {@code
 * OverflowAwareTextFitterTest#fitOverflowingText_shrinksOnlyUntilCollisionClears_whenFullBoxFitIsUnreachable}
 * - pour rester deterministes quelle que soit la police effectivement
 * resolue sur la machine d'execution.
 */
class NeighborShapeOverlapFixerTest {

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

    /**
     * Meme technique que {@code OverflowAwareTextFitterTest} : l'API XSLF publique
     * ne propose aucun setter pour le nom d'une forme, on doit donc passer par le
     * XML sous-jacent - necessaire ici pour construire un titre reconnu par {@link
     * TitleDetector} (prefixe de nom "Titre"/"Title"/...) dans les tests de la
     * section "Formes exemptees de l'elargissement sommaire".
     */
    private static void setShapeName(XSLFShape shape, String name) {
        Object xmlObject = shape.getXmlObject();
        if (xmlObject instanceof CTShape) {
            ((CTShape) xmlObject).getNvSpPr().getCNvPr().setName(name);
        }
    }

    @Test
    void computeOccupiedZone_topAlignment_startsAtAnchorTop() {
        Rectangle2D anchor = new Rectangle2D.Double(0, 100, 200, 300);
        Rectangle2D occupied = NeighborShapeOverlapFixer.computeOccupiedZone(anchor, 50, VerticalAlignment.TOP);

        assertEquals(100, occupied.getY(), 0.001);
        assertEquals(50, occupied.getHeight(), 0.001);
    }

    @Test
    void computeOccupiedZone_bottomAlignment_endsAtAnchorBottom() {
        Rectangle2D anchor = new Rectangle2D.Double(0, 100, 200, 300);
        Rectangle2D occupied = NeighborShapeOverlapFixer.computeOccupiedZone(anchor, 50, VerticalAlignment.BOTTOM);

        assertEquals(350, occupied.getY(), 0.001, "doit se terminer au bas de l'ancre (100+300-50)");
        assertEquals(50, occupied.getHeight(), 0.001);
    }

    @Test
    void computeOccupiedZone_middleAlignment_isCentered() {
        Rectangle2D anchor = new Rectangle2D.Double(0, 100, 200, 300);
        Rectangle2D occupied = NeighborShapeOverlapFixer.computeOccupiedZone(anchor, 100, VerticalAlignment.MIDDLE);

        assertEquals(200, occupied.getY(), 0.001, "100 + (300-100)/2");
        assertEquals(100, occupied.getHeight(), 0.001);
    }

    @Test
    void findCollidingShape_detectsOverlapWithNeighbourTextShape() {
        XSLFTextBox self = textBox(0, 100, 200, 10, "Forme dont on mesure le texte");
        XSLFTextBox neighbour = textBox(0, 110, 200, 50, "Forme voisine avec du texte visible");
        List<XSLFShape> all = List.of(self, neighbour);

        Rectangle2D occupied = new Rectangle2D.Double(0, 100, 200, 20);

        assertEquals(neighbour, NeighborShapeOverlapFixer.findCollidingShape(occupied, self, all));
    }

    @Test
    void findCollidingShape_ignoresShapesWithoutText() {
        XSLFTextBox self = textBox(0, 100, 200, 10, "Texte");
        XSLFTextBox emptyPanel = textBox(0, 100, 200, 50, ""); // panneau de fond, sans texte
        List<XSLFShape> all = List.of(self, emptyPanel);

        Rectangle2D occupied = new Rectangle2D.Double(0, 100, 200, 20);

        assertNull(NeighborShapeOverlapFixer.findCollidingShape(occupied, self, all),
                "une forme sans texte (fond/panneau) ne doit jamais etre consideree comme une collision");
    }

    @Test
    void findCollidingShape_returnsNullWhenIsolated() {
        XSLFTextBox self = textBox(0, 100, 200, 10, "Texte isole");
        XSLFTextBox farAway = textBox(500, 500, 50, 50, "Loin de toute zone occupee");
        List<XSLFShape> all = List.of(self, farAway);

        Rectangle2D occupied = new Rectangle2D.Double(0, 100, 200, 20);

        assertNull(NeighborShapeOverlapFixer.findCollidingShape(occupied, self, all));
    }

    @Test
    void findCollidingShape_ignoresCalloutNeighbour() {
        // Voir Javadoc de la classe, section "Exemption des bulles narratives
        // (callouts) de tout chevauchement" : motif reel du slide 4 ("Title 3" vs
        // "Bulle narrative : rectangle a coins arrondis 1") - l'ancre d'une bulle de
        // rappel englobe geometriquement son bec, jusqu'a l'endroit qu'il designe ; ce
        // "chevauchement" purement geometrique ne doit plus jamais etre remonte comme
        // une collision, meme quand l'ancre intersecte reellement la zone occupee.
        XSLFTextBox self = textBox(0, 100, 200, 10, "Forme dont on mesure le texte");
        XSLFTextBox callout = textBox(0, 110, 200, 50, "~4 ou 35 ? dossiers par departement");
        callout.setShapeType(ShapeType.WEDGE_ROUND_RECT_CALLOUT);
        List<XSLFShape> all = List.of(self, callout);

        Rectangle2D occupied = new Rectangle2D.Double(0, 100, 200, 20);

        assertNull(NeighborShapeOverlapFixer.findCollidingShape(occupied, self, all),
                "une bulle narrative (callout) ne doit jamais etre retenue comme forme voisine 'colliding'");
    }

    @Test
    void fixNeighborOverlaps_neverCorrectsCalloutShape_evenWhenItCollidesWithNeighbour() throws IOException {
        // La bulle elle-meme ne doit jamais etre corrigee a cause d'une collision avec
        // une forme voisine (elle est replacee au premier plan a la fin de la
        // composition d'origine - voir Javadoc de la classe).
        XSLFTextBox callout = textBox(0, 100, 200, 50, "~4 ou 35 ? dossiers par departement");
        callout.setShapeType(ShapeType.WEDGE_ROUND_RECT_CALLOUT);
        callout.getTextParagraphs().get(0).setLineSpacing(100.0);
        XSLFTextBox neighbour = textBox(0, 100, 200, 200, "Forme voisine largement chevauchee");

        BufferedImage img = new BufferedImage(10, 10, BufferedImage.TYPE_INT_ARGB);
        Graphics2D graphics = img.createGraphics();
        try {
            Rectangle2D calloutAnchor = callout.getAnchor();
            double calloutTextHeight = callout.getTextHeight(graphics);
            Rectangle2D calloutOccupied = NeighborShapeOverlapFixer.computeOccupiedZone(
                    calloutAnchor, calloutTextHeight, callout.getVerticalAlignment());
            assertTrue(calloutOccupied.intersects(neighbour.getAnchor()),
                    "precondition : la bulle doit reellement chevaucher la forme voisine");

            int changed = NeighborShapeOverlapFixer.fixNeighborOverlaps(slide, graphics);

            assertEquals(0, changed, "une bulle narrative ne doit jamais etre corrigee, meme en cas de collision reelle");
            assertEquals(100.0, callout.getTextParagraphs().get(0).getLineSpacing(), 0.001,
                    "l'interligne de la bulle ne doit jamais avoir ete touche");
        } finally {
            graphics.dispose();
        }
    }

    @Test
    void fixNeighborOverlaps_neverShrinksNeighbour_becauseOfCalloutOverlap() throws IOException {
        // Motif reel du slide 4 : "Title 3" ne doit plus jamais etre retreci a cause
        // d'un chevauchement mesure avec l'ancre (bec compris) de la bulle narrative.
        Rectangle2D anchor = new Rectangle2D.Double(0, 100, 300, 200);
        XSLFTextBox title = textBox(anchor.getX(), anchor.getY(), anchor.getWidth(), anchor.getHeight(),
                "Texte de Title 3, largement dans sa propre ancre");
        title.getTextParagraphs().get(0).setLineSpacing(100.0);
        title.getTextParagraphs().get(0).getTextRuns().get(0).setFontSize(18.0);

        // Ancre de la bulle demarrant au meme point que celle de Title 3 (comme le bec
        // d'une bulle qui pointe vers le haut du bloc de texte) : garantit une
        // intersection avec la zone occupee quelle que soit la hauteur de texte
        // reellement mesuree sur la machine d'execution.
        XSLFTextBox callout = textBox(anchor.getX(), anchor.getY(), anchor.getWidth(), 100, "~4 ou 35 ? dossiers par departement");
        callout.setShapeType(ShapeType.WEDGE_ROUND_RECT_CALLOUT);

        BufferedImage img = new BufferedImage(10, 10, BufferedImage.TYPE_INT_ARGB);
        Graphics2D graphics = img.createGraphics();
        try {
            double textHeight = title.getTextHeight(graphics);
            Rectangle2D occupied = NeighborShapeOverlapFixer.computeOccupiedZone(anchor, textHeight, title.getVerticalAlignment());
            assertTrue(occupied.intersects(callout.getAnchor()),
                    "precondition : l'ancre de Title 3 doit reellement chevaucher l'ancre (bec compris) de la bulle");

            int changed = NeighborShapeOverlapFixer.fixNeighborOverlaps(slide, graphics);

            assertEquals(0, changed, "un chevauchement avec une bulle narrative ne doit jamais declencher de correction");
            assertEquals(100.0, title.getTextParagraphs().get(0).getLineSpacing(), 0.001,
                    "l'interligne de la forme voisine de la bulle ne doit jamais avoir ete touche");
            assertEquals(18.0, title.getTextParagraphs().get(0).getTextRuns().get(0).getFontSize(), 0.001,
                    "la police de la forme voisine de la bulle ne doit jamais avoir ete touchee");
        } finally {
            graphics.dispose();
        }
    }

    @Test
    void fixNeighborOverlaps_skipsShape_whenTextExceedsOwnAnchor() throws IOException {
        // Meme technique que OverflowAwareTextFitterTest : anchor minuscule (5pt) et
        // police large (60pt), garantit un texte qui deborde de sa PROPRE ancre quelle
        // que soit la police resolue sur la machine d'execution. Ce cas reste
        // entierement du ressort d'OverflowAwareTextFitter (qui doit s'executer avant) -
        // NeighborShapeOverlapFixer ne doit jamais y toucher, meme en presence d'une
        // collision reelle avec la forme voisine.
        XSLFTextBox box = textBox(10, 20, 300, 5, "Texte qui necessite bien plus de place que 5pt de haut");
        box.getTextParagraphs().get(0).getTextRuns().get(0).setFontSize(60.0);
        XSLFTextBox neighbour = textBox(10, 25, 300, 200, "Forme voisine avec du texte visible");

        BufferedImage img = new BufferedImage(10, 10, BufferedImage.TYPE_INT_ARGB);
        Graphics2D graphics = img.createGraphics();
        try {
            Rectangle2D anchor = box.getAnchor();
            double measuredTextHeight = box.getTextHeight(graphics);
            assertTrue(measuredTextHeight > anchor.getHeight(), "precondition : le texte doit deborder de sa propre ancre");

            int changed = NeighborShapeOverlapFixer.fixNeighborOverlaps(slide, graphics);

            assertEquals(0, changed);
            assertEquals(60.0, box.getTextParagraphs().get(0).getTextRuns().get(0).getFontSize(), 0.001);
        } finally {
            graphics.dispose();
        }
    }

    @Test
    void fixNeighborOverlaps_resolvesCollision_forSommaireBroadeningExemptShape_evenWhenTextExceedsOwnAnchor()
            throws IOException {
        // Complement du test precedent (voir Javadoc de la classe, "Formes exemptees
        // de l'elargissement sommaire", 2026-08-31) : motif reel du slide 2 (sommaire)
        // du fichier "Refonte BEL" - un item de sommaire, autofit non declare
        // localement (aucun appel a setTextAutofit, motif "tout XSLFTextBox
        // fraichement cree"), sur une ancre dimensionnee pour une seule ligne alors
        // que son texte reel en occupe plusieurs a cette largeur. OverflowAwareTextFitter
        // exempte desormais deliberement une telle forme de tout retrecissement (voir
        // sa Javadoc) - ce correctif doit prendre le relais si ca cree une collision
        // avec une forme voisine, CONTRAIREMENT au test precedent (forme non exemptee,
        // toujours ignoree dans ce cas).
        XSLFTextBox title = textBox(0, -200, 300, 100, "Sommaire");
        setShapeName(title, "Titre 1");

        Rectangle2D anchor = new Rectangle2D.Double(0, 100, 120, 20);
        XSLFTextBox item = slide.createTextBox();
        item.setAnchor(anchor);
        item.getTextParagraphs().get(0).getTextRuns().get(0).setText(
                "Un texte assez long pour deborder sur plusieurs lignes a cette largeur");
        item.getTextParagraphs().get(0).getTextRuns().get(0).setFontSize(20.0);
        item.getTextParagraphs().get(0).setLineSpacing(100.0);

        BufferedImage img = new BufferedImage(10, 10, BufferedImage.TYPE_INT_ARGB);
        Graphics2D graphics = img.createGraphics();
        try {
            double fullHeight = item.getTextHeight(graphics);
            assertTrue(fullHeight > anchor.getHeight(),
                    "precondition : le texte doit deborder de sa PROPRE ancre (motif vise par ce test)");

            item.getTextParagraphs().get(0).setLineSpacing(100.0 * 0.85);
            item.getTextParagraphs().get(0).getTextRuns().get(0).setFontSize(20.0 * 0.25);
            double minPossibleHeight = item.getTextHeight(graphics);

            item.getTextParagraphs().get(0).setLineSpacing(100.0);
            item.getTextParagraphs().get(0).getTextRuns().get(0).setFontSize(20.0);

            assertTrue(minPossibleHeight < fullHeight,
                    "precondition : interligne+police reduits doivent mesurer moins que l'etat d'origine");

            double margin = 3.0;
            double neighbourTop = anchor.getY() + minPossibleHeight + margin;
            XSLFTextBox neighbour = textBox(0, neighbourTop, 120, 200, "Item suivant du sommaire");

            Rectangle2D fullOccupied = NeighborShapeOverlapFixer.computeOccupiedZone(
                    anchor, fullHeight, item.getVerticalAlignment());
            assertTrue(fullOccupied.intersects(neighbour.getAnchor()),
                    "precondition : a taille pleine, le texte doit chevaucher la forme voisine");

            int changed = NeighborShapeOverlapFixer.fixNeighborOverlaps(slide, graphics);

            assertEquals(1, changed, "une forme exemptee par l'elargissement sommaire doit etre prise en charge "
                    + "par ce correctif, meme quand son texte depasse sa propre ancre");
            double heightAfter = item.getTextHeight(graphics);
            Rectangle2D occupiedAfter = NeighborShapeOverlapFixer.computeOccupiedZone(
                    anchor, heightAfter, item.getVerticalAlignment());
            assertFalse(occupiedAfter.intersects(neighbour.getAnchor()),
                    "le chevauchement avec l'item suivant doit avoir disparu apres correction");
        } finally {
            graphics.dispose();
        }
    }

    @Test
    void fixNeighborOverlaps_doesNothing_whenIsolated() throws IOException {
        XSLFTextBox box = textBox(0, 100, 300, 500, "Texte normal, largement dans sa boite");
        textBox(0, 1_000_000, 300, 50, "Forme voisine tres loin, jamais atteinte");

        BufferedImage img = new BufferedImage(10, 10, BufferedImage.TYPE_INT_ARGB);
        Graphics2D graphics = img.createGraphics();
        try {
            int changed = NeighborShapeOverlapFixer.fixNeighborOverlaps(slide, graphics);
            assertEquals(0, changed);
        } finally {
            graphics.dispose();
        }
    }

    /** Construit une boite a plusieurs paragraphes courts, tous a interligne en pourcentage (non declare). */
    private XSLFTextBox multiLineBox(Rectangle2D anchor, int lineCount, double fontSize) {
        XSLFTextBox box = slide.createTextBox();
        box.setAnchor(anchor);
        box.setTextAutofit(TextShape.TextAutofit.NONE);
        box.getTextParagraphs().get(0).getTextRuns().get(0).setText("Ligne 0");
        box.getTextParagraphs().get(0).getTextRuns().get(0).setFontSize(fontSize);
        for (int i = 1; i < lineCount; i++) {
            XSLFTextParagraph para = box.addNewTextParagraph();
            XSLFTextRun run = para.addNewTextRun();
            run.setText("Ligne " + i);
            run.setFontSize(fontSize);
        }
        return box;
    }

    private void setAllFontSizes(XSLFTextBox box, double size) {
        for (XSLFTextParagraph para : box.getTextParagraphs()) {
            for (XSLFTextRun run : para.getTextRuns()) {
                run.setFontSize(size);
            }
        }
    }

    private void setAllLineSpacing(XSLFTextBox box, double pct) {
        for (XSLFTextParagraph para : box.getTextParagraphs()) {
            para.setLineSpacing(pct);
        }
    }

    @Test
    void fixNeighborOverlaps_resolvesCollision_whenNeighbourWithinOwnGenerousAnchor() throws IOException {
        // Reproduit le motif reel (slide 10, boite "44 %..." vs annotation "Nos
        // échanges..." - voir Javadoc de la classe) : le texte tient largement dans sa
        // propre ancre (jamais de debordement au sens d'OverflowAwareTextFitter), mais
        // une forme voisine independante, positionnee dans l'espace laisse vide sous le
        // texte, se retrouve neanmoins chevauchee par le texte reellement mesure.
        //
        // La position de la forme voisine est calibree a partir de DEUX mesures reelles
        // (taille d'origine, puis taille au maximum de ce que NeighborShapeOverlapFixer
        // peut atteindre en combinant interligne a 85% et police a 25%) plutot que de
        // coordonnees fixes, pour rester deterministe quelle que soit la police
        // effectivement resolue sur la machine d'execution.
        Rectangle2D anchor = new Rectangle2D.Double(0, 100, 300, 2000);
        XSLFTextBox box = multiLineBox(anchor, 10, 20.0);
        setAllLineSpacing(box, 100.0);

        BufferedImage img = new BufferedImage(10, 10, BufferedImage.TYPE_INT_ARGB);
        Graphics2D graphics = img.createGraphics();
        try {
            double fullHeight = box.getTextHeight(graphics);

            setAllLineSpacing(box, 100.0 * 0.85);
            setAllFontSizes(box, 20.0 * 0.25);
            double minPossibleHeight = box.getTextHeight(graphics);

            // Restaure l'etat d'origine avant l'appel reel au correctif.
            setAllLineSpacing(box, 100.0);
            setAllFontSizes(box, 20.0);

            assertTrue(minPossibleHeight < fullHeight,
                    "precondition : interligne+police reduits doivent mesurer moins que l'etat d'origine");

            double margin = 3.0;
            double neighbourTop = anchor.getY() + minPossibleHeight + margin;
            double neighbourHeight = (anchor.getY() + anchor.getHeight()) - neighbourTop;
            assertTrue(neighbourHeight > 0, "precondition de calibrage : l'ancre genereuse doit laisser de la place sous la forme voisine");
            XSLFTextBox neighbour = textBox(0, neighbourTop, 300, neighbourHeight, "Forme voisine independante");

            Rectangle2D fullOccupied = NeighborShapeOverlapFixer.computeOccupiedZone(anchor, fullHeight, box.getVerticalAlignment());
            assertTrue(fullOccupied.intersects(neighbour.getAnchor()),
                    "precondition : le texte a sa taille d'origine doit bien chevaucher la forme voisine");
            assertTrue(fullHeight <= anchor.getHeight(), "precondition du perimetre de ce correctif : pas de debordement de la propre ancre");

            int changed = NeighborShapeOverlapFixer.fixNeighborOverlaps(slide, graphics);

            assertEquals(1, changed);
            double heightAfter = box.getTextHeight(graphics);
            Rectangle2D occupiedAfter = NeighborShapeOverlapFixer.computeOccupiedZone(anchor, heightAfter, box.getVerticalAlignment());
            assertFalse(occupiedAfter.intersects(neighbour.getAnchor()),
                    "le chevauchement avec la forme voisine doit avoir disparu apres correction");
        } finally {
            graphics.dispose();
        }
    }

    @Test
    void fixNeighborOverlaps_restoresOriginalValues_whenCollisionCannotBeResolved() throws IOException {
        // Forme voisine positionnee des le sommet de l'ancre (comme l'ancre elle-meme,
        // alignement TOP par defaut) : la zone occupee du texte commence toujours a
        // anchor.getY(), quelle que soit la reduction d'interligne ou de police
        // appliquee - aucune combinaison des deux leviers ne peut donc jamais faire
        // disparaitre cette collision. Le correctif doit alors renoncer et restaurer
        // exactement les valeurs d'origine plutot que de livrer un texte ecrase pour
        // rien.
        Rectangle2D anchor = new Rectangle2D.Double(0, 100, 300, 2000);
        XSLFTextBox box = multiLineBox(anchor, 5, 20.0);
        setAllLineSpacing(box, 100.0);
        XSLFTextBox neighbour = textBox(0, anchor.getY(), 300, 2000, "Forme voisine couvrant toute l'ancre");

        BufferedImage img = new BufferedImage(10, 10, BufferedImage.TYPE_INT_ARGB);
        Graphics2D graphics = img.createGraphics();
        try {
            double originalHeight = box.getTextHeight(graphics);
            assertTrue(originalHeight <= anchor.getHeight(), "precondition du perimetre de ce correctif");
            Rectangle2D originalOccupied = NeighborShapeOverlapFixer.computeOccupiedZone(anchor, originalHeight, box.getVerticalAlignment());
            assertTrue(originalOccupied.intersects(neighbour.getAnchor()), "precondition : collision reelle attendue");

            int changed = NeighborShapeOverlapFixer.fixNeighborOverlaps(slide, graphics);

            assertEquals(0, changed, "une collision insoluble ne doit jamais etre comptee comme corrigee");
            for (XSLFTextParagraph para : box.getTextParagraphs()) {
                assertEquals(100.0, para.getLineSpacing(), 0.001, "l'interligne d'origine doit etre restaure");
                for (XSLFTextRun run : para.getTextRuns()) {
                    assertEquals(20.0, run.getFontSize(), 0.001, "la taille de police d'origine doit etre restauree");
                }
            }
        } finally {
            graphics.dispose();
        }
    }

    /** Ajoute un paragraphe ENTIEREMENT VIDE (espaceur : endParaRPr seul, aucun {@code <a:r>}) - motif du slide 17. */
    private XSLFTextParagraph addBlankSpacerParagraph(XSLFTextBox box, double lineSpacing) {
        XSLFTextParagraph blank = box.addNewTextParagraph();
        blank.setLineSpacing(lineSpacing);
        return blank;
    }

    @Test
    void fixNeighborOverlaps_ignoresCollision_whenOnlyBlankSpacerParagraphsOverlapNeighbour() throws IOException {
        // Reproduit le motif reel du slide 17 (voir Javadoc de la classe, "2e variante") :
        // une forme voisine independante est positionnee exactement la ou seuls des
        // paragraphes ENTIEREMENT VIDES (espaceurs) atterrissent. Le bloc ENTIER (espaceurs
        // compris) chevauche bien cette forme voisine - un correctif "forme entiere" y
        // verrait un probleme - mais AUCUN paragraphe visible n'est reellement concerne :
        // rien ne doit etre modifie.
        Rectangle2D anchor = new Rectangle2D.Double(0, 100, 300, 2000);
        XSLFTextBox box = slide.createTextBox();
        box.setAnchor(anchor);
        box.setTextAutofit(TextShape.TextAutofit.NONE);
        box.getTextParagraphs().get(0).getTextRuns().get(0).setText("Texte visible");
        box.getTextParagraphs().get(0).getTextRuns().get(0).setFontSize(20.0);
        box.getTextParagraphs().get(0).setLineSpacing(100.0);

        BufferedImage img = new BufferedImage(10, 10, BufferedImage.TYPE_INT_ARGB);
        Graphics2D graphics = img.createGraphics();
        try {
            // Mesure du seul paragraphe visible AVANT d'ajouter les espaceurs, pour caler
            // la forme voisine juste derriere lui - dans la zone que seuls des espaceurs
            // occuperont ensuite.
            double soleParagraphHeight = box.getTextHeight(graphics);

            for (int i = 0; i < 3; i++) {
                addBlankSpacerParagraph(box, 100.0);
            }

            double fullHeight = box.getTextHeight(graphics);
            assertTrue(fullHeight > soleParagraphHeight,
                    "precondition : les espaceurs vides doivent contribuer une hauteur mesuree non nulle");

            double margin = 2.0;
            double neighbourTop = anchor.getY() + soleParagraphHeight + margin;
            XSLFTextBox neighbour = textBox(0, neighbourTop, 300, 500, "Voisin dans la zone des espaceurs");

            Rectangle2D wholeOccupied = NeighborShapeOverlapFixer.computeOccupiedZone(anchor, fullHeight, box.getVerticalAlignment());
            assertTrue(wholeOccupied.intersects(neighbour.getAnchor()),
                    "precondition : le bloc ENTIER (espaceurs compris) doit chevaucher la forme voisine");

            int changed = NeighborShapeOverlapFixer.fixNeighborOverlaps(slide, graphics);

            assertEquals(0, changed, "aucun paragraphe visible ne chevauche reellement ce voisin (seuls des "
                    + "espaceurs vides s'y trouvent) : rien ne doit etre corrige");
            for (XSLFTextParagraph para : box.getTextParagraphs()) {
                assertEquals(100.0, para.getLineSpacing(), 0.001, "l'interligne ne doit jamais avoir ete touche");
            }
            assertEquals(20.0, box.getTextParagraphs().get(0).getTextRuns().get(0).getFontSize(), 0.001,
                    "la police du paragraphe visible ne doit jamais avoir ete touchee");
        } finally {
            graphics.dispose();
        }
    }

    @Test
    void fixNeighborOverlaps_resolvesCollision_whenVisibleParagraphAfterBlankSpacersOverlaps() throws IOException {
        // Toujours le motif du slide 17, mais cette fois un DERNIER paragraphe visible,
        // place apres les espaceurs, chevauche bien reellement la forme voisine - le
        // correctif doit alors agir (interligne puis, en repli, police - inchange), en
        // s'arretant des que CE paragraphe cesse individuellement de chevaucher.
        Rectangle2D anchor = new Rectangle2D.Double(0, 100, 300, 2000);
        XSLFTextBox box = slide.createTextBox();
        box.setAnchor(anchor);
        box.setTextAutofit(TextShape.TextAutofit.NONE);
        box.getTextParagraphs().get(0).getTextRuns().get(0).setText("Texte visible");
        box.getTextParagraphs().get(0).getTextRuns().get(0).setFontSize(20.0);
        box.getTextParagraphs().get(0).setLineSpacing(100.0);
        for (int i = 0; i < 5; i++) {
            addBlankSpacerParagraph(box, 100.0);
        }
        XSLFTextParagraph target = box.addNewTextParagraph();
        target.setLineSpacing(100.0);
        XSLFTextRun targetRun = target.addNewTextRun();
        targetRun.setText("Paragraphe cible");
        targetRun.setFontSize(20.0);

        BufferedImage img = new BufferedImage(10, 10, BufferedImage.TYPE_INT_ARGB);
        Graphics2D graphics = img.createGraphics();
        try {
            double fullHeight = box.getTextHeight(graphics);

            setAllLineSpacing(box, 100.0 * 0.85);
            setAllFontSizes(box, 20.0 * 0.25);
            double minPossibleHeight = box.getTextHeight(graphics);

            setAllLineSpacing(box, 100.0);
            setAllFontSizes(box, 20.0);

            assertTrue(minPossibleHeight < fullHeight,
                    "precondition : interligne+police reduits doivent mesurer moins que l'etat d'origine");

            double margin = 3.0;
            double neighbourTop = anchor.getY() + minPossibleHeight + margin;
            double neighbourHeight = (anchor.getY() + anchor.getHeight()) - neighbourTop;
            assertTrue(neighbourHeight > 0, "precondition de calibrage : l'ancre genereuse doit laisser de la place sous la forme voisine");
            XSLFTextBox neighbour = textBox(0, neighbourTop, 300, neighbourHeight, "Forme voisine independante");

            int changed = NeighborShapeOverlapFixer.fixNeighborOverlaps(slide, graphics);

            assertEquals(1, changed);
            double heightAfter = box.getTextHeight(graphics);
            Rectangle2D occupiedAfter = NeighborShapeOverlapFixer.computeOccupiedZone(anchor, heightAfter, box.getVerticalAlignment());
            assertFalse(occupiedAfter.intersects(neighbour.getAnchor()),
                    "le chevauchement avec la forme voisine doit avoir disparu apres correction");
        } finally {
            graphics.dispose();
        }
    }
}
