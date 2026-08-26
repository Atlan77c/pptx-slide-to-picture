package io.github.atlan77c.pptx2picture.internal;

import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertSame;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.awt.Color;
import java.awt.Graphics2D;
import java.awt.image.BufferedImage;

import org.apache.poi.sl.draw.Drawable;
import org.apache.poi.sl.usermodel.ShapeType;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFConnectorShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.drawingml.x2006.main.CTLineEndProperties;
import org.openxmlformats.schemas.drawingml.x2006.main.CTLineProperties;
import org.openxmlformats.schemas.drawingml.x2006.main.STLineEndType;
import org.openxmlformats.schemas.presentationml.x2006.main.CTConnector;

/**
 * Tests unitaires pour {@link ConnectorArrowFixer}.
 *
 * <p>Ces tests verifient le comportement de {@code installBeforeDraw}/
 * {@code restoreAfterDraw} depuis le passage a l'architecture v13
 * (substitution via {@code DrawFactory}, voir Javadoc de {@link ConnectorArrowFixer}) :
 * contrairement aux versions v11/v12, le connecteur n'est plus jamais retire de
 * la slide - c'est POI lui-meme qui le dessine, via notre {@code Drawable} de
 * remplacement, exactement a sa place naturelle dans l'ordre d'empilement.
 */
class ConnectorArrowFixerTest {

    @Test
    void installBeforeDraw_thenSlideDraw_doesNotThrow_forBentConnectorWithDeclaredArrow() throws Exception {
        try (XMLSlideShow ppt = new XMLSlideShow()) {
            XSLFSlide slide = ppt.createSlide();
            XSLFConnectorShape connector = slide.createConnector();
            connector.setShapeType(ShapeType.BENT_CONNECTOR_3);
            connector.setAnchor(new java.awt.geom.Rectangle2D.Double(10, 10, 100, 60));
            connector.setLineWidth(2.0);
            connector.setLineColor(Color.BLACK);
            setTailArrow(connector);

            BufferedImage img = new BufferedImage(400, 400, BufferedImage.TYPE_INT_ARGB);
            Graphics2D graphics = img.createGraphics();
            try {
                graphics.setColor(Color.WHITE);
                graphics.fillRect(0, 0, 400, 400);

                Object previous = ConnectorArrowFixer.installBeforeDraw(graphics);
                slide.draw(graphics);
                ConnectorArrowFixer.restoreAfterDraw(graphics, previous);

                assertTrue(slide.getShapes().contains(connector),
                        "le connecteur ne doit jamais etre retire de la slide (architecture v13)");
                assertTrue(hasNonWhitePixel(img), "quelque chose doit avoir ete dessine");
            } finally {
                graphics.dispose();
            }
        }
    }

    @Test
    void installBeforeDraw_thenSlideDraw_doesNotThrow_forCurvedConnectorWithDeclaredArrow() throws Exception {
        try (XMLSlideShow ppt = new XMLSlideShow()) {
            XSLFSlide slide = ppt.createSlide();
            XSLFConnectorShape connector = slide.createConnector();
            connector.setShapeType(ShapeType.CURVED_CONNECTOR_3);
            connector.setAnchor(new java.awt.geom.Rectangle2D.Double(10, 10, 100, 60));
            connector.setLineWidth(2.0);
            connector.setLineColor(Color.BLACK);
            setHeadArrow(connector);

            BufferedImage img = new BufferedImage(400, 400, BufferedImage.TYPE_INT_ARGB);
            Graphics2D graphics = img.createGraphics();
            try {
                graphics.setColor(Color.WHITE);
                graphics.fillRect(0, 0, 400, 400);

                Object previous = ConnectorArrowFixer.installBeforeDraw(graphics);
                slide.draw(graphics);
                ConnectorArrowFixer.restoreAfterDraw(graphics, previous);

                assertTrue(slide.getShapes().contains(connector),
                        "le connecteur ne doit jamais etre retire de la slide (architecture v13)");
                assertTrue(hasNonWhitePixel(img), "quelque chose doit avoir ete dessine");
            } finally {
                graphics.dispose();
            }
        }
    }

    @Test
    void installBeforeDraw_thenSlideDraw_leavesStraightConnectorUntouched() throws Exception {
        try (XMLSlideShow ppt = new XMLSlideShow()) {
            XSLFSlide slide = ppt.createSlide();
            XSLFConnectorShape connector = slide.createConnector();
            connector.setShapeType(ShapeType.STRAIGHT_CONNECTOR_1);
            connector.setAnchor(new java.awt.geom.Rectangle2D.Double(10, 10, 100, 60));
            connector.setLineWidth(2.0);
            connector.setLineColor(Color.BLACK);
            setTailArrow(connector);

            BufferedImage img = new BufferedImage(400, 400, BufferedImage.TYPE_INT_ARGB);
            Graphics2D graphics = img.createGraphics();
            try {
                Object previous = ConnectorArrowFixer.installBeforeDraw(graphics);
                slide.draw(graphics);
                ConnectorArrowFixer.restoreAfterDraw(graphics, previous);

                assertTrue(slide.getShapes().contains(connector), "un connecteur droit ne doit jamais etre modifie");
                assertTrue(getLn(connector).isSetTailEnd(), "la decoration d'origine ne doit pas etre touchee");
            } finally {
                graphics.dispose();
            }
        }
    }

    @Test
    void installBeforeDraw_thenSlideDraw_leavesConnectorWithoutArrowUntouched() throws Exception {
        try (XMLSlideShow ppt = new XMLSlideShow()) {
            XSLFSlide slide = ppt.createSlide();
            XSLFConnectorShape connector = slide.createConnector();
            connector.setShapeType(ShapeType.BENT_CONNECTOR_3);
            connector.setAnchor(new java.awt.geom.Rectangle2D.Double(10, 10, 100, 60));
            connector.setLineWidth(2.0);
            connector.setLineColor(Color.BLACK);
            // Pas de pointe de fleche declaree.

            BufferedImage img = new BufferedImage(400, 400, BufferedImage.TYPE_INT_ARGB);
            Graphics2D graphics = img.createGraphics();
            try {
                Object previous = ConnectorArrowFixer.installBeforeDraw(graphics);
                slide.draw(graphics);
                ConnectorArrowFixer.restoreAfterDraw(graphics, previous);

                assertTrue(slide.getShapes().contains(connector),
                        "rien a corriger sur un connecteur sans pointe de fleche declaree");
            } finally {
                graphics.dispose();
            }
        }
    }

    @Test
    void installBeforeDraw_thenRestoreAfterDraw_roundTripsThePreviousDrawFactoryHint() {
        BufferedImage img = new BufferedImage(10, 10, BufferedImage.TYPE_INT_ARGB);
        Graphics2D graphics = img.createGraphics();
        try {
            assertNull(graphics.getRenderingHint(Drawable.DRAW_FACTORY),
                    "aucun DrawFactory personnalise n'est installe au depart dans ce test");

            Object previous = ConnectorArrowFixer.installBeforeDraw(graphics);
            Object installed = graphics.getRenderingHint(Drawable.DRAW_FACTORY);
            assertTrue(installed != null, "installBeforeDraw doit avoir installe un DrawFactory");

            ConnectorArrowFixer.restoreAfterDraw(graphics, previous);
            assertSame(previous, graphics.getRenderingHint(Drawable.DRAW_FACTORY),
                    "restoreAfterDraw doit remettre exactement le hint precedent");
        } finally {
            graphics.dispose();
        }
    }

    private static boolean hasNonWhitePixel(BufferedImage img) {
        for (int y = 0; y < img.getHeight(); y++) {
            for (int x = 0; x < img.getWidth(); x++) {
                if ((img.getRGB(x, y) & 0x00FFFFFF) != 0x00FFFFFF) {
                    return true;
                }
            }
        }
        return false;
    }

    private static void setTailArrow(XSLFConnectorShape connector) {
        CTLineEndProperties tailEnd = getOrCreateTailEnd(connector);
        tailEnd.setType(STLineEndType.TRIANGLE);
        tailEnd.setW(org.openxmlformats.schemas.drawingml.x2006.main.STLineEndWidth.MED);
        tailEnd.setLen(org.openxmlformats.schemas.drawingml.x2006.main.STLineEndLength.MED);
    }

    private static void setHeadArrow(XSLFConnectorShape connector) {
        CTLineEndProperties headEnd = getOrCreateHeadEnd(connector);
        headEnd.setType(STLineEndType.TRIANGLE);
        headEnd.setW(org.openxmlformats.schemas.drawingml.x2006.main.STLineEndWidth.MED);
        headEnd.setLen(org.openxmlformats.schemas.drawingml.x2006.main.STLineEndLength.MED);
    }

    private static CTLineProperties getLn(XSLFConnectorShape connector) {
        CTConnector ct = (CTConnector) connector.getXmlObject();
        return ct.getSpPr().getLn();
    }

    private static CTLineEndProperties getOrCreateTailEnd(XSLFConnectorShape connector) {
        CTLineProperties ln = getLn(connector);
        return ln.isSetTailEnd() ? ln.getTailEnd() : ln.addNewTailEnd();
    }

    private static CTLineEndProperties getOrCreateHeadEnd(XSLFConnectorShape connector) {
        CTLineProperties ln = getLn(connector);
        return ln.isSetHeadEnd() ? ln.getHeadEnd() : ln.addNewHeadEnd();
    }
}
