package io.github.atlan77c.pptx2picture.internal;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertSame;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.awt.Color;
import java.awt.Graphics2D;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.IOException;

import javax.imageio.ImageIO;

import org.apache.poi.sl.draw.Drawable;
import org.apache.poi.sl.usermodel.PictureData.PictureType;
import org.apache.poi.sl.usermodel.ShapeType;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFPictureData;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.junit.jupiter.api.Test;

/**
 * Tests unitaires pour {@link PictureGeometryClipFixer}.
 *
 * <p>Reproduit le bug d'origine (image "decoupee selon une forme" dans
 * PowerPoint mais rendue integralement rectangulaire par POI - voir Javadoc
 * de {@link PictureGeometryClipFixer}) avec une image de couleur unie
 * inscrite dans une geometrie {@link ShapeType#ELLIPSE} : sans le correctif,
 * un pixel proche d'un coin de la boite englobante de l'image (hors de
 * l'ellipse inscrite, mais a l'interieur du rectangle d'ancrage) porterait la
 * couleur de l'image au lieu du fond - exactement le symptome observe
 * (photo qui deborde de la forme).
 */
class PictureGeometryClipFixerTest {

    private static final Color PICTURE_COLOR = Color.RED;

    @Test
    void installBeforeDraw_thenSlideDraw_clipsPictureToEllipseGeometry() throws Exception {
        try (XMLSlideShow ppt = new XMLSlideShow()) {
            XSLFSlide slide = ppt.createSlide();
            XSLFPictureData pictureData = ppt.addPicture(solidColorPng(100, 100, PICTURE_COLOR), PictureType.PNG);
            XSLFPictureShape pic = slide.createPicture(pictureData);
            pic.setAnchor(new Rectangle2D.Double(50, 50, 200, 200));
            pic.setShapeType(ShapeType.ELLIPSE);

            BufferedImage img = new BufferedImage(400, 400, BufferedImage.TYPE_INT_ARGB);
            Graphics2D graphics = img.createGraphics();
            try {
                graphics.setColor(Color.WHITE);
                graphics.fillRect(0, 0, 400, 400);

                Object previous = PictureGeometryClipFixer.installBeforeDraw(graphics);
                slide.draw(graphics);
                PictureGeometryClipFixer.restoreAfterDraw(graphics, previous);

                // Ancre : (50,50) a (250,250), ellipse inscrite centree en (150,150), rayons 100/100.
                // (55,55) est loin a l'interieur du rectangle d'ancrage mais nettement hors de
                // l'ellipse ((dx/100)^2+(dy/100)^2 = 1.805 > 1) : doit rester le fond si le
                // correctif fonctionne - c'est exactement le symptome du bug d'origine sinon.
                assertEquals(rgb(Color.WHITE), rgb(img.getRGB(55, 55)),
                        "un coin de l'ancre, hors de l'ellipse inscrite, doit rester le fond (image decoupee)");
                // Le centre de l'ellipse doit en revanche bien afficher la couleur de la photo.
                assertEquals(rgb(PICTURE_COLOR), rgb(img.getRGB(150, 150)),
                        "le centre de l'ellipse doit afficher la couleur de la photo");
            } finally {
                graphics.dispose();
            }
        }
    }

    @Test
    void installBeforeDraw_thenSlideDraw_leavesRectangularPictureUnclipped() throws Exception {
        try (XMLSlideShow ppt = new XMLSlideShow()) {
            XSLFSlide slide = ppt.createSlide();
            XSLFPictureData pictureData = ppt.addPicture(solidColorPng(100, 100, PICTURE_COLOR), PictureType.PNG);
            XSLFPictureShape pic = slide.createPicture(pictureData);
            pic.setAnchor(new Rectangle2D.Double(50, 50, 200, 200));
            // ShapeType.RECT (par defaut pour une image inseree normalement) : aucune decoupe
            // ne doit s'appliquer, voir PictureGeometryClipFixer.ResolvedGeometry#qualifiesForClipping.

            BufferedImage img = new BufferedImage(400, 400, BufferedImage.TYPE_INT_ARGB);
            Graphics2D graphics = img.createGraphics();
            try {
                graphics.setColor(Color.WHITE);
                graphics.fillRect(0, 0, 400, 400);

                Object previous = PictureGeometryClipFixer.installBeforeDraw(graphics);
                slide.draw(graphics);
                PictureGeometryClipFixer.restoreAfterDraw(graphics, previous);

                assertEquals(rgb(PICTURE_COLOR), rgb(img.getRGB(55, 55)),
                        "une image rectangulaire normale ne doit jamais etre decoupee (aucune regression)");
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

            Object previous = PictureGeometryClipFixer.installBeforeDraw(graphics);
            Object installed = graphics.getRenderingHint(Drawable.DRAW_FACTORY);
            assertTrue(installed != null, "installBeforeDraw doit avoir installe un DrawFactory");

            PictureGeometryClipFixer.restoreAfterDraw(graphics, previous);
            assertSame(previous, graphics.getRenderingHint(Drawable.DRAW_FACTORY),
                    "restoreAfterDraw doit remettre exactement le hint precedent");
        } finally {
            graphics.dispose();
        }
    }

    private static int rgb(int argb) {
        return argb & 0x00FFFFFF;
    }

    private static int rgb(Color color) {
        return color.getRGB() & 0x00FFFFFF;
    }

    private static byte[] solidColorPng(int width, int height, Color color) throws IOException {
        BufferedImage img = new BufferedImage(width, height, BufferedImage.TYPE_INT_RGB);
        Graphics2D g = img.createGraphics();
        try {
            g.setColor(color);
            g.fillRect(0, 0, width, height);
        } finally {
            g.dispose();
        }
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        ImageIO.write(img, "png", baos);
        return baos.toByteArray();
    }
}
