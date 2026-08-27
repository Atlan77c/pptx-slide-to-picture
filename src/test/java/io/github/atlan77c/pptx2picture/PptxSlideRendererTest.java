package io.github.atlan77c.pptx2picture;

import org.apache.poi.sl.usermodel.PictureData.PictureType;
import org.apache.poi.sl.usermodel.ShapeType;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFPictureData;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextBox;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import javax.imageio.ImageIO;
import java.awt.Color;
import java.awt.Dimension;
import java.awt.Graphics2D;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Path;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * Tests d'integration de l'API publique, sur des fichiers pptx generes
 * programmatiquement (pas de fichier .pptx du monde reel embarque dans le
 * depot : le contenu de test doit rester libre de tout droit et independant
 * de toute police non standard).
 */
class PptxSlideRendererTest {

    @TempDir
    Path tempDir;

    private File buildPptx(int slideCount) throws IOException {
        try (XMLSlideShow ppt = new XMLSlideShow()) {
            for (int i = 0; i < slideCount; i++) {
                XSLFSlide slide = ppt.createSlide();
                XSLFTextBox box = slide.createTextBox();
                box.setAnchor(new Rectangle2D.Double(50, 50, 400, 100));
                box.setText("Slide " + (i + 1));
            }
            File file = tempDir.resolve("test-" + slideCount + "-slides.pptx").toFile();
            try (FileOutputStream out = new FileOutputStream(file)) {
                ppt.write(out);
            }
            return file;
        }
    }

    private Dimension pageSizeOf(File pptxFile) throws IOException {
        try (XMLSlideShow ppt = new XMLSlideShow(new java.io.FileInputStream(pptxFile))) {
            return ppt.getPageSize();
        }
    }

    @Test
    void renderSlide_producesImageScaledToRequestedFactor() throws Exception {
        File pptx = buildPptx(1);
        Dimension pageSize = pageSizeOf(pptx);

        RenderOptions options = RenderOptions.builder().scale(1.5f).build();
        BufferedImage image = PptxSlideRenderer.renderSlide(pptx, 1, options);

        assertEquals(Math.round(pageSize.width * 1.5f), image.getWidth());
        assertEquals(Math.round(pageSize.height * 1.5f), image.getHeight());
    }

    @Test
    void renderSlide_defaultOptions_usesScaleOfTwo() throws Exception {
        File pptx = buildPptx(1);
        Dimension pageSize = pageSizeOf(pptx);

        BufferedImage image = PptxSlideRenderer.renderSlide(pptx, 1);

        assertEquals(Math.round(pageSize.width * 2.0f), image.getWidth());
        assertEquals(Math.round(pageSize.height * 2.0f), image.getHeight());
    }

    @Test
    void renderSlide_indexTooLow_throwsPptxRenderException() throws Exception {
        File pptx = buildPptx(2);

        PptxRenderException ex = assertThrows(PptxRenderException.class,
                () -> PptxSlideRenderer.renderSlide(pptx, 0));
        assertTrue(ex.getMessage().contains("Index de slide invalide"));
    }

    @Test
    void renderSlide_indexTooHigh_throwsPptxRenderException() throws Exception {
        File pptx = buildPptx(2);

        PptxRenderException ex = assertThrows(PptxRenderException.class,
                () -> PptxSlideRenderer.renderSlide(pptx, 3));
        assertTrue(ex.getMessage().contains("Index de slide invalide"));
    }

    @Test
    void renderSlideToFile_writesReadablePng() throws Exception {
        File pptx = buildPptx(1);
        Dimension pageSize = pageSizeOf(pptx);
        File output = tempDir.resolve("out.png").toFile();

        PptxSlideRenderer.renderSlideToFile(pptx, 1, output);

        assertTrue(output.isFile());
        BufferedImage readBack = ImageIO.read(output);
        assertNotNull(readBack);
        assertEquals(Math.round(pageSize.width * 2.0f), readBack.getWidth());
        assertEquals(Math.round(pageSize.height * 2.0f), readBack.getHeight());
    }

    @Test
    void getSlideCount_returnsNumberOfSlides() throws Exception {
        File pptx = buildPptx(3);
        assertEquals(3, PptxSlideRenderer.getSlideCount(pptx));
    }

    @Test
    void renderSlide_withFixTextOverflowDisabled_stillRendersWithoutError() throws Exception {
        File pptx = buildPptx(1);
        RenderOptions options = RenderOptions.builder().fixTextOverflow(false).build();

        BufferedImage image = PptxSlideRenderer.renderSlide(pptx, 1, options);

        assertNotNull(image);
    }

    @Test
    void renderSlide_onNonExistentFile_throwsPptxRenderException() {
        File missing = tempDir.resolve("does-not-exist.pptx").toFile();
        assertThrows(PptxRenderException.class, () -> PptxSlideRenderer.renderSlide(missing, 1));
    }

    @Test
    void renderSlideToFile_jpegFormat_writesReadableOpaqueJpeg() throws Exception {
        File pptx = buildPptx(1);
        Dimension pageSize = pageSizeOf(pptx);
        File output = tempDir.resolve("out.jpg").toFile();
        RenderOptions options = RenderOptions.builder().format(OutputFormat.JPEG).build();

        PptxSlideRenderer.renderSlideToFile(pptx, 1, output, options);

        assertTrue(output.isFile());
        BufferedImage readBack = ImageIO.read(output);
        assertNotNull(readBack);
        assertEquals(Math.round(pageSize.width * 2.0f), readBack.getWidth());
        assertEquals(Math.round(pageSize.height * 2.0f), readBack.getHeight());
        // Le JPEG ne supporte pas la transparence : le type lu ne doit jamais porter de canal alpha.
        assertTrue(readBack.getColorModel().getNumComponents() == 3
                || !readBack.getColorModel().hasAlpha());
    }

    @Test
    void renderSlideToFile_jpegFormat_lowQualityProducesSmallerFileThanHighQuality() throws Exception {
        File pptx = buildPptx(1);
        File lowQualityOutput = tempDir.resolve("low.jpg").toFile();
        File highQualityOutput = tempDir.resolve("high.jpg").toFile();

        PptxSlideRenderer.renderSlideToFile(pptx, 1, lowQualityOutput,
                RenderOptions.builder().format(OutputFormat.JPEG).jpegQuality(0.05f).build());
        PptxSlideRenderer.renderSlideToFile(pptx, 1, highQualityOutput,
                RenderOptions.builder().format(OutputFormat.JPEG).jpegQuality(1.0f).build());

        assertTrue(lowQualityOutput.length() < highQualityOutput.length());
    }

    @Test
    void renderSlideAsSvg_producesWellFormedSvgContainingText() throws Exception {
        File pptx = buildPptx(1);

        String svg = PptxSlideRenderer.renderSlideAsSvg(pptx, 1, RenderOptions.defaults());

        assertNotNull(svg);
        assertTrue(svg.contains("<svg"));
        assertTrue(svg.contains("Slide 1"));
    }

    @Test
    void renderSlideToFile_svgFormat_writesFileStartingWithXmlDeclaration() throws Exception {
        File pptx = buildPptx(1);
        File output = tempDir.resolve("out.svg").toFile();
        RenderOptions options = RenderOptions.builder().format(OutputFormat.SVG).build();

        PptxSlideRenderer.renderSlideToFile(pptx, 1, output, options);

        assertTrue(output.isFile());
        String content = java.nio.file.Files.readString(output.toPath());
        assertTrue(content.contains("<svg"));
    }

    /**
     * Test de bout en bout, via l'API publique complete (donc y compris
     * {@code DrawFactoryComposer}, pas seulement le fixer isole comme dans
     * {@code PictureGeometryClipFixerTest}), du bug d'origine : une image
     * "decoupee selon une forme" dans PowerPoint (ici une ellipse) qui
     * ressortait integralement rectangulaire au rendu, recouvrant tout ce qui
     * se trouve derriere elle dans la boite englobante de son ancre.
     */
    @Test
    void renderSlide_clipsPictureCroppedToEllipseShape() throws Exception {
        Color pictureColor = Color.RED;
        File pptx;
        try (XMLSlideShow ppt = new XMLSlideShow()) {
            XSLFSlide slide = ppt.createSlide();
            XSLFPictureData pictureData = ppt.addPicture(solidColorPng(100, 100, pictureColor), PictureType.PNG);
            XSLFPictureShape pic = slide.createPicture(pictureData);
            pic.setAnchor(new Rectangle2D.Double(50, 50, 200, 200));
            pic.setShapeType(ShapeType.ELLIPSE);

            pptx = tempDir.resolve("cropped-to-ellipse.pptx").toFile();
            try (FileOutputStream out = new FileOutputStream(pptx)) {
                ppt.write(out);
            }
        }

        // scale=1 pour que les coordonnees pixel de l'image produite correspondent
        // directement aux coordonnees (en points) de la slide.
        RenderOptions options = RenderOptions.builder().scale(1.0f).build();
        BufferedImage image = PptxSlideRenderer.renderSlide(pptx, 1, options);

        // Ancre : (50,50) a (250,250), ellipse inscrite centree en (150,150), rayons 100/100.
        // (55,55) est loin a l'interieur du rectangle d'ancrage mais nettement hors de
        // l'ellipse ((dx/100)^2+(dy/100)^2 = 1.805 > 1) : doit rester le fond (blanc, la
        // valeur par defaut de RenderOptions) si la decoupe est appliquee de bout en bout -
        // c'est exactement le symptome du bug d'origine sinon (photo qui deborde de la forme).
        assertEquals(rgb(Color.WHITE), rgb(image.getRGB(55, 55)),
                "un coin de l'ancre, hors de l'ellipse inscrite, doit rester le fond");
        assertEquals(rgb(pictureColor), rgb(image.getRGB(150, 150)),
                "le centre de l'ellipse doit afficher la couleur de la photo");
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
