package io.github.atlan77c.pptx2picture.internal;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFGroupShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTable;
import org.apache.poi.xslf.usermodel.XSLFTableCell;
import org.apache.poi.xslf.usermodel.XSLFTableRow;
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
import static org.junit.jupiter.api.Assertions.assertSame;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * Reproduit le motif reel a l'origine de {@link TableCellLineSpacingFixer} (slide 18,
 * "Tableau 6", "Mes Evenements Emploi et Prestation - Doc vision 1.0.pptx") - voir Javadoc
 * de la classe testee. Comme {@code OverflowAwareTextFitterTest} (meme convention, meme
 * raison), les assertions evitent toute comparaison numerique de hauteur/interligne
 * dependante des metriques Java2D/AWT reelles (qui varient selon les polices installees
 * sur la machine d'execution) - la correction de fidelite elle-meme est verifiee a la
 * place sur un rendu reel, cote utilisateur.
 */
class TableCellLineSpacingFixerTest {

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

    /** Table 1x1 avec une geometrie volontairement genereuse (aucun debordement possible en phase B). */
    private XSLFTableCell singleCell(String text, double lineSpacing) {
        XSLFTable table = slide.createTable(1, 1);
        table.setAnchor(new Rectangle2D.Double(0, 0, 300, 500));
        table.setColumnWidth(0, 300);
        XSLFTableRow row = table.getRows().get(0);
        row.setHeight(500); // tres genereux : le texte court utilise ici ne peut jamais deborder
        XSLFTableCell cell = row.getCells().get(0);
        cell.setText(text);
        if (lineSpacing != 0) {
            cell.getTextParagraphs().get(0).setLineSpacing(lineSpacing);
        }
        return cell;
    }

    @Test
    void fixTableCellLineSpacing_leavesLineSpacingUnchanged_whenNoneDeclaredLocally() {
        // Motif reel de la cellule diagnostiquee slide 18 : aucun <a:pPr> local, donc
        // aucun heritage resolu par une forme toute neuve sans theme/masque (meme
        // constat que OverflowAwareTextFitterTest#correctPercentLineSpacingForFidelity_
        // leavesUnchanged_whenNoLineSpacingDeclared) - getLineSpacing() reste null, ce
        // garde-fou de OverflowAwareTextFitter (reutilise tel quel) ne s'applique donc
        // pas ici non plus.
        XSLFTableCell cell = singleCell("Texte sans interligne declare localement", 0);

        BufferedImage img = new BufferedImage(10, 10, BufferedImage.TYPE_INT_ARGB);
        Graphics2D graphics = img.createGraphics();
        try {
            int shrunk = TableCellLineSpacingFixer.fixTableCellLineSpacing(slide, graphics);

            assertEquals(0, shrunk, "geometrie genereuse : aucun retrecissement de secours necessaire");
            assertEquals(null, cell.getTextParagraphs().get(0).getLineSpacing());
        } finally {
            graphics.dispose();
        }
    }

    @Test
    void fixTableCellLineSpacing_leavesLineSpacingUnchanged_whenSpacingIsAbsolutePoints() {
        // Convention POI : une valeur NEGATIVE signifie un interligne ABSOLU en points
        // (spcPts) - hors de portee du correctif de fidelite reutilise (voir Javadoc de
        // OverflowAwareTextFitter#correctPercentLineSpacingForFidelity).
        XSLFTableCell cell = singleCell("Texte a interligne absolu", -24.0);

        BufferedImage img = new BufferedImage(10, 10, BufferedImage.TYPE_INT_ARGB);
        Graphics2D graphics = img.createGraphics();
        try {
            int shrunk = TableCellLineSpacingFixer.fixTableCellLineSpacing(slide, graphics);

            assertEquals(0, shrunk);
            assertEquals(-24.0, cell.getTextParagraphs().get(0).getLineSpacing(), 0.001,
                    "un interligne absolu en points ne doit jamais etre modifie par ce correctif");
        } finally {
            graphics.dispose();
        }
    }

    @Test
    void fixTableCellLineSpacing_returnsZero_whenSlideHasNoTable() {
        // Ne doit jamais se declencher ni lever d'exception sur un slide sans tableau.
        XSLFTextBox box = slide.createTextBox();
        box.setText("Pas de tableau ici");

        BufferedImage img = new BufferedImage(10, 10, BufferedImage.TYPE_INT_ARGB);
        Graphics2D graphics = img.createGraphics();
        try {
            int shrunk = TableCellLineSpacingFixer.fixTableCellLineSpacing(slide, graphics);
            assertEquals(0, shrunk);
        } finally {
            graphics.dispose();
        }
    }

    @Test
    void shrinkIfStillOverflowing_shrinksFont_whenTextOverflowsCellAnchor() {
        // Meme technique que OverflowAwareTextFitterTest#fitOverflowingText_shapeAutofit_
        // shrinksFontInsteadOfGrowingAnchor : ancre volontairement minuscule (5pt) et
        // police large (60pt), garantit un debordement massif quelle que soit la police
        // reellement resolue sur la machine d'execution. Teste directement (visibilite
        // package) plutot qu'a travers fixTableCellLineSpacing/updateCellAnchor : ce
        // dernier fait grandir dynamiquement la hauteur de ligne selon le texte mesure
        // (voir Javadoc de la classe), donc une geometrie de table "normale" ne peut
        // quasiment jamais reproduire un vrai debordement de la cellule au-dela de sa
        // PROPRE ancre - seul ce retrecissement de secours, isole, s'y prete.
        XSLFTable table = slide.createTable(1, 1);
        table.setAnchor(new Rectangle2D.Double(0, 0, 300, 5));
        table.setColumnWidth(0, 300);
        XSLFTableRow row = table.getRows().get(0);
        row.setHeight(5);
        XSLFTableCell cell = row.getCells().get(0);
        cell.setText("Texte qui necessite bien plus de place que 5pt de haut");
        cell.getTextParagraphs().get(0).getTextRuns().get(0).setFontSize(60.0);
        Rectangle2D forcedAnchor = new Rectangle2D.Double(0, 0, 300, 5);
        cell.setAnchor(forcedAnchor);

        BufferedImage img = new BufferedImage(10, 10, BufferedImage.TYPE_INT_ARGB);
        Graphics2D graphics = img.createGraphics();
        try {
            double measuredTextHeight = cell.getTextHeight(graphics);
            assertTrue(measuredTextHeight > forcedAnchor.getHeight(),
                    "precondition du test : le texte doit deborder de l'ancre forcee");

            boolean shrunk = TableCellLineSpacingFixer.shrinkIfStillOverflowing(cell, graphics);

            assertTrue(shrunk);
            Rectangle2D anchorAfter = cell.getAnchor();
            assertEquals(forcedAnchor.getWidth(), anchorAfter.getWidth(), 0.001, "la largeur ne doit pas changer");
            assertEquals(forcedAnchor.getHeight(), anchorAfter.getHeight(), 0.001,
                    "la hauteur de l'ancre ne doit pas changer - seule la police est retrecie");
            Double fontSizeAfter = cell.getTextParagraphs().get(0).getTextRuns().get(0).getFontSize();
            assertTrue(fontSizeAfter < 60.0, "la police doit avoir ete retrecie sous sa taille d'origine (60pt)");
        } finally {
            graphics.dispose();
        }
    }

    @Test
    void shrinkIfStillOverflowing_doesNothing_whenTextFitsWithinAnchor() {
        XSLFTable table = slide.createTable(1, 1);
        table.setAnchor(new Rectangle2D.Double(0, 0, 300, 500));
        table.setColumnWidth(0, 300);
        XSLFTableRow row = table.getRows().get(0);
        row.setHeight(500);
        XSLFTableCell cell = row.getCells().get(0);
        cell.setText("Texte court");
        cell.setAnchor(new Rectangle2D.Double(0, 0, 300, 500));

        BufferedImage img = new BufferedImage(10, 10, BufferedImage.TYPE_INT_ARGB);
        Graphics2D graphics = img.createGraphics();
        try {
            boolean shrunk = TableCellLineSpacingFixer.shrinkIfStillOverflowing(cell, graphics);
            assertEquals(false, shrunk);
        } finally {
            graphics.dispose();
        }
    }

    @Test
    void collectTables_findsTopLevelTable() {
        XSLFTable table = slide.createTable(1, 1);
        table.setAnchor(new Rectangle2D.Double(0, 0, 100, 50));
        table.setColumnWidth(0, 100);
        table.getRows().get(0).setHeight(50);

        List<XSLFTable> found = TableCellLineSpacingFixer.collectTables(slide.getShapes());

        assertEquals(1, found.size());
        assertSame(table, found.get(0));
    }

    @Test
    void collectTables_descendsIntoGroups() {
        XSLFGroupShape group = slide.createGroup();
        group.setAnchor(new Rectangle2D.Double(0, 0, 200, 200));
        group.setInteriorAnchor(new Rectangle2D.Double(0, 0, 200, 200));
        XSLFTable table = group.createTable(1, 1);
        table.setAnchor(new Rectangle2D.Double(0, 0, 100, 50));
        table.setColumnWidth(0, 100);
        table.getRows().get(0).setHeight(50);

        // Le tableau n'est PAS un enfant direct du slide (uniquement du groupe) : ne doit
        // etre trouve que via la recursion dans XSLFGroupShape#getShapes().
        List<XSLFTable> found = TableCellLineSpacingFixer.collectTables(slide.getShapes());

        assertEquals(1, found.size());
        assertSame(table, found.get(0));
    }

    @Test
    void collectCellsWithVisibleText_excludesBlankAndEmptyCells() {
        XSLFTable table = slide.createTable(1, 3);
        table.setAnchor(new Rectangle2D.Double(0, 0, 300, 50));
        table.setColumnWidth(0, 100);
        table.setColumnWidth(1, 100);
        table.setColumnWidth(2, 100);
        table.getRows().get(0).setHeight(50);
        List<XSLFTableCell> cells = table.getRows().get(0).getCells();
        cells.get(0).setText("Visible");
        cells.get(1).setText("   "); // blanc uniquement
        cells.get(2).setText(""); // vide

        List<XSLFTableCell> found = TableCellLineSpacingFixer.collectCellsWithVisibleText(table);

        assertEquals(1, found.size());
        assertSame(cells.get(0), found.get(0));
    }
}
