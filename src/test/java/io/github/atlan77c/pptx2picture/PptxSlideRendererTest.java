package io.github.atlan77c.pptx2picture;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextBox;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import javax.imageio.ImageIO;
import java.awt.Dimension;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
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
}
