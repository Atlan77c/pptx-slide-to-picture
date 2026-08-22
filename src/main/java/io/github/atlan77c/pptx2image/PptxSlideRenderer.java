package io.github.atlan77c.pptx2image;

import io.github.atlan77c.pptx2image.internal.OverflowAwareTextFitter;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.imageio.ImageIO;
import java.awt.Dimension;
import java.awt.Graphics2D;
import java.awt.RenderingHints;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;
import java.util.Objects;

/**
 * Convertit une slide d'un fichier PowerPoint ({@code .pptx}) en image
 * ({@link BufferedImage} ou PNG), en pur Java via Apache POI - sans
 * dependance a LibreOffice, a PowerPoint (automatisation COM) ni a un
 * navigateur headless.
 *
 * <h2>Exemple d'utilisation</h2>
 * <pre>{@code
 * BufferedImage image = PptxSlideRenderer.renderSlide(new File("presentation.pptx"), 4);
 * ImageIO.write(image, "png", new File("slide4.png"));
 *
 * // ou directement :
 * PptxSlideRenderer.renderSlideToFile(new File("presentation.pptx"), 4, new File("slide4.png"));
 * }</pre>
 *
 * <h2>Fidelite du rendu et limites connues</h2>
 * <p>Le texte, les formes vectorielles, les images et les tableaux natifs
 * sont rendus fidelement. Deux limites sont documentees (voir
 * {@code conversion_pptx_vers_images.md} du projet d'origine pour le detail
 * complet de l'investigation) :
 * <ul>
 *   <li><b>Les graphiques integres (Excel/{@code XSLFChart}) ne sont pas rendus</b>
 *       par Apache POI - un espace vide apparait a leur emplacement. Contournement
 *       possible non implemente ici : extraire les donnees via l'API {@code XSLFChart}/
 *       {@code XDDFChart} et dessiner le graphique separement (ex. avec JFreeChart).</li>
 *   <li><b>Ecart de metriques de police</b> : pour certaines polices (observe avec
 *       "PoliceX", la police de l'Etat francais), Java2D/AWT surestime la hauteur de
 *       texte necessaire d'environ 30 a 35% par rapport au moteur de rendu natif de
 *       PowerPoint, ce qui peut faire deborder une zone de texte hors de sa boite.
 *       Corrige par defaut (voir {@link RenderOptions#isFixTextOverflow()}) en
 *       retrecissant la police des formes concernees, mais seulement lorsque ce
 *       debordement chevaucherait reellement une autre forme de texte voisine -
 *       desactivable via {@link RenderOptions.Builder#fixTextOverflow(boolean)}.</li>
 * </ul>
 */
public final class PptxSlideRenderer {

    private static final Logger LOG = LoggerFactory.getLogger(PptxSlideRenderer.class);

    private PptxSlideRenderer() {
    }

    /** Equivalent a {@link #renderSlide(File, int, RenderOptions)} avec {@link RenderOptions#defaults()}. */
    public static BufferedImage renderSlide(File pptxFile, int slideIndex) throws PptxRenderException {
        return renderSlide(pptxFile, slideIndex, RenderOptions.defaults());
    }

    /**
     * Rend une slide en image.
     *
     * @param pptxFile   le fichier {@code .pptx} source.
     * @param slideIndex l'index de la slide a rendre, en base 1 (la premiere
     *                    slide du fichier est l'index 1, pas 0).
     * @param options    les options de rendu (echelle, correctif de
     *                    debordement, couleur de fond).
     * @return l'image de la slide rendue.
     * @throws PptxRenderException si le fichier est illisible/invalide, si
     *                              {@code slideIndex} est hors bornes, ou en
     *                              cas d'erreur interne de rendu.
     */
    public static BufferedImage renderSlide(File pptxFile, int slideIndex, RenderOptions options) throws PptxRenderException {
        Objects.requireNonNull(pptxFile, "pptxFile");
        Objects.requireNonNull(options, "options");

        try (InputStream fis = new FileInputStream(pptxFile);
             XMLSlideShow ppt = new XMLSlideShow(fis)) {

            List<XSLFSlide> slides = ppt.getSlides();
            if (slideIndex < 1 || slideIndex > slides.size()) {
                throw new PptxRenderException(String.format(
                        "Index de slide invalide: %d (le fichier contient %d slide(s), index attendu entre 1 et %d)",
                        slideIndex, slides.size(), slides.size()));
            }

            Dimension pageSize = ppt.getPageSize();
            float scale = options.getScale();
            int width = Math.round(pageSize.width * scale);
            int height = Math.round(pageSize.height * scale);

            BufferedImage img = new BufferedImage(width, height, BufferedImage.TYPE_INT_ARGB);
            Graphics2D graphics = img.createGraphics();
            try {
                graphics.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
                graphics.setRenderingHint(RenderingHints.KEY_TEXT_ANTIALIASING, RenderingHints.VALUE_TEXT_ANTIALIAS_ON);
                graphics.setRenderingHint(RenderingHints.KEY_INTERPOLATION, RenderingHints.VALUE_INTERPOLATION_BICUBIC);

                graphics.setPaint(options.getBackground());
                graphics.fill(new Rectangle2D.Float(0, 0, width, height));
                graphics.scale(scale, scale);
                graphics.fill(new Rectangle2D.Float(0, 0, pageSize.width, pageSize.height));

                XSLFSlide slide = slides.get(slideIndex - 1);

                if (options.isFixTextOverflow()) {
                    int shrunk = OverflowAwareTextFitter.fitOverflowingText(slide, graphics);
                    if (shrunk > 0 && LOG.isDebugEnabled()) {
                        LOG.debug("{} forme(s) de texte retrecie(s) pour corriger un debordement (slide {})", shrunk, slideIndex);
                    }
                }

                slide.draw(graphics);
            } finally {
                graphics.dispose();
            }

            return img;
        } catch (PptxRenderException e) {
            throw e;
        } catch (IOException e) {
            throw new PptxRenderException("Impossible de lire le fichier pptx: " + pptxFile, e);
        } catch (RuntimeException e) {
            // Apache POI (parsing OOXML, rendu Graphics2D) leve parfois des exceptions
            // non verifiees (format invalide, forme non supportee...) : on les
            // enveloppe pour garder une API d'erreur unique et previsible.
            throw new PptxRenderException("Erreur lors du rendu de la slide " + slideIndex + " de " + pptxFile, e);
        }
    }

    /** Equivalent a {@link #renderSlideToFile(File, int, File, RenderOptions)} avec {@link RenderOptions#defaults()}. */
    public static void renderSlideToFile(File pptxFile, int slideIndex, File outputPngFile) throws PptxRenderException {
        renderSlideToFile(pptxFile, slideIndex, outputPngFile, RenderOptions.defaults());
    }

    /**
     * Rend une slide et ecrit directement le resultat au format PNG dans
     * {@code outputPngFile}. Equivalent a appeler {@link #renderSlide} puis
     * {@link ImageIO#write(java.awt.image.RenderedImage, String, File)}.
     *
     * @throws PptxRenderException si le rendu echoue, ou si l'ecriture du
     *                              fichier PNG de sortie echoue.
     */
    public static void renderSlideToFile(File pptxFile, int slideIndex, File outputPngFile, RenderOptions options) throws PptxRenderException {
        Objects.requireNonNull(outputPngFile, "outputPngFile");
        BufferedImage image = renderSlide(pptxFile, slideIndex, options);
        try (FileOutputStream out = new FileOutputStream(outputPngFile)) {
            ImageIO.write(image, "png", out);
        } catch (IOException e) {
            throw new PptxRenderException("Impossible d'ecrire l'image PNG de sortie: " + outputPngFile, e);
        }
    }

    /**
     * Retourne le nombre de slides du fichier pptx, utile pour valider un
     * {@code slideIndex} ou parcourir toutes les slides d'un fichier.
     *
     * @throws PptxRenderException si le fichier est illisible/invalide.
     */
    public static int getSlideCount(File pptxFile) throws PptxRenderException {
        Objects.requireNonNull(pptxFile, "pptxFile");
        try (InputStream fis = new FileInputStream(pptxFile);
             XMLSlideShow ppt = new XMLSlideShow(fis)) {
            return ppt.getSlides().size();
        } catch (IOException e) {
            throw new PptxRenderException("Impossible de lire le fichier pptx: " + pptxFile, e);
        }
    }
}
