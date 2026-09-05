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
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFPictureData;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.presentationml.x2006.main.CTPicture;

/**
 * Tests unitaires pour {@link PictureAlphaModFixer}.
 *
 * <p>Reproduit le bug d'origine (image rendue partiellement transparente dans
 * PowerPoint via {@code <a:alphaModFix>}, mais rendue entierement opaque par
 * POI - voir Javadoc de {@link PictureAlphaModFixer}) avec une image de
 * couleur unie posee sur un fond de couleur differente : sans le correctif,
 * un pixel au centre de l'image porterait la couleur pleine de l'image au
 * lieu d'un melange avec le fond - exactement le symptome observe (image
 * semi-transparente affichee opaque).
 */
class PictureAlphaModFixerTest {

    private static final Color BACKGROUND_COLOR = Color.WHITE;
    private static final Color PICTURE_COLOR = Color.RED;

    @Test
    void installBeforeDraw_thenSlideDraw_blendsPictureWithAlphaModFixTransparency() throws Exception {
        try (XMLSlideShow ppt = new XMLSlideShow()) {
            XSLFSlide slide = ppt.createSlide();
            XSLFPictureData pictureData = ppt.addPicture(solidColorPng(100, 100, PICTURE_COLOR), PictureType.PNG);
            XSLFPictureShape pic = slide.createPicture(pictureData);
            pic.setAnchor(new Rectangle2D.Double(50, 50, 200, 200));
            setAlphaModFix(pic, 50_000); // 50% d'opacite, soit 50% de transparence dans PowerPoint.

            BufferedImage img = new BufferedImage(400, 400, BufferedImage.TYPE_INT_ARGB);
            Graphics2D graphics = img.createGraphics();
            try {
                graphics.setColor(BACKGROUND_COLOR);
                graphics.fillRect(0, 0, 400, 400);

                Object previous = PictureAlphaModFixer.installBeforeDraw(graphics);
                slide.draw(graphics);
                PictureAlphaModFixer.restoreAfterDraw(graphics, previous);

                // Centre de l'image (150,150) : melange 50/50 rouge/blanc attendu - ni la
                // couleur pleine de l'image (bug d'origine) ni le fond pur.
                int blended = img.getRGB(150, 150);
                assertTrue(rgb(blended) != rgb(PICTURE_COLOR),
                        "l'image ne doit plus etre rendue pleinement opaque (bug d'origine)");
                assertTrue(rgb(blended) != rgb(BACKGROUND_COLOR),
                        "l'image doit rester visible, pas totalement transparente");
                int green = (blended >> 8) & 0xFF;
                int blue = blended & 0xFF;
                assertTrue(green > 90 && green < 165, "canal vert melange ~50/50 attendu, obtenu " + green);
                assertTrue(blue > 90 && blue < 165, "canal bleu melange ~50/50 attendu, obtenu " + blue);
            } finally {
                graphics.dispose();
            }
        }
    }

    @Test
    void installBeforeDraw_thenSlideDraw_leavesFullyOpaquePictureUnchanged() throws Exception {
        try (XMLSlideShow ppt = new XMLSlideShow()) {
            XSLFSlide slide = ppt.createSlide();
            XSLFPictureData pictureData = ppt.addPicture(solidColorPng(100, 100, PICTURE_COLOR), PictureType.PNG);
            XSLFPictureShape pic = slide.createPicture(pictureData);
            pic.setAnchor(new Rectangle2D.Double(50, 50, 200, 200));
            // Aucun <a:alphaModFix> declare : comportement standard de POI attendu, sans
            // la moindre regression (voir PictureAlphaModFixer.resolveAlphaModFixAmt).

            BufferedImage img = new BufferedImage(400, 400, BufferedImage.TYPE_INT_ARGB);
            Graphics2D graphics = img.createGraphics();
            try {
                graphics.setColor(BACKGROUND_COLOR);
                graphics.fillRect(0, 0, 400, 400);

                Object previous = PictureAlphaModFixer.installBeforeDraw(graphics);
                slide.draw(graphics);
                PictureAlphaModFixer.restoreAfterDraw(graphics, previous);

                assertEquals(rgb(PICTURE_COLOR), rgb(img.getRGB(150, 150)),
                        "une image sans alphaModFix ne doit jamais etre rendue partiellement transparente "
                                + "(aucune regression)");
            } finally {
                graphics.dispose();
            }
        }
    }

    @Test
    void installBeforeDraw_thenSlideDraw_treatsExplicitFullyOpaqueAmountAsUnchanged() throws Exception {
        try (XMLSlideShow ppt = new XMLSlideShow()) {
            XSLFSlide slide = ppt.createSlide();
            XSLFPictureData pictureData = ppt.addPicture(solidColorPng(100, 100, PICTURE_COLOR), PictureType.PNG);
            XSLFPictureShape pic = slide.createPicture(pictureData);
            pic.setAnchor(new Rectangle2D.Double(50, 50, 200, 200));
            setAlphaModFix(pic, 100_000); // amt=100000 declare explicitement : deja pleinement opaque.

            BufferedImage img = new BufferedImage(400, 400, BufferedImage.TYPE_INT_ARGB);
            Graphics2D graphics = img.createGraphics();
            try {
                graphics.setColor(BACKGROUND_COLOR);
                graphics.fillRect(0, 0, 400, 400);

                Object previous = PictureAlphaModFixer.installBeforeDraw(graphics);
                slide.draw(graphics);
                PictureAlphaModFixer.restoreAfterDraw(graphics, previous);

                assertEquals(rgb(PICTURE_COLOR), rgb(img.getRGB(150, 150)),
                        "amt=100000 (pleinement opaque) ne doit entrainer aucun changement de rendu");
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

            Object previous = PictureAlphaModFixer.installBeforeDraw(graphics);
            Object installed = graphics.getRenderingHint(Drawable.DRAW_FACTORY);
            assertTrue(installed != null, "installBeforeDraw doit avoir installe un DrawFactory");

            PictureAlphaModFixer.restoreAfterDraw(graphics, previous);
            assertSame(previous, graphics.getRenderingHint(Drawable.DRAW_FACTORY),
                    "restoreAfterDraw doit remettre exactement le hint precedent");
        } finally {
            graphics.dispose();
        }
    }

    /** Declare {@code <a:alphaModFix amt="..."/>} sur le {@code <a:blip>} de {@code pic} - pas d'API publique dediee cote Apache POI 5.2.5 (voir Javadoc de {@link PictureAlphaModFixer}). */
    private static void setAlphaModFix(XSLFPictureShape pic, int amt) {
        CTPicture ctPicture = (CTPicture) pic.getXmlObject();
        ctPicture.getBlipFill().getBlip().addNewAlphaModFix().setAmt(amt);
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
