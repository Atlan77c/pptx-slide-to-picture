package io.github.atlan77c.pptx2picture;

import io.github.atlan77c.pptx2picture.internal.OverflowAwareTextFitter;
import io.github.atlan77c.pptx2picture.internal.SymbolFontRunFixer;
import org.apache.batik.dom.GenericDOMImplementation;
import org.apache.batik.svggen.SVGGraphics2D;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.w3c.dom.DOMImplementation;
import org.w3c.dom.Document;

import javax.imageio.IIOImage;
import javax.imageio.ImageIO;
import javax.imageio.ImageWriteParam;
import javax.imageio.ImageWriter;
import javax.imageio.stream.ImageOutputStream;
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
import java.io.OutputStreamWriter;
import java.io.Writer;
import java.nio.charset.StandardCharsets;
import java.util.Iterator;
import java.util.List;
import java.util.Objects;

/**
 * Convertit une slide d'un fichier PowerPoint ({@code .pptx}) en image
 * ({@link BufferedImage}, PNG, JPEG ou SVG), en pur Java via Apache POI (rendu
 * raster) et Apache Batik (rendu vectoriel) - sans dependance a LibreOffice, a
 * PowerPoint (automatisation COM) ni a un navigateur headless.
 *
 * <h2>Exemple d'utilisation</h2>
 * <pre>{@code
 * BufferedImage image = PptxSlideRenderer.renderSlide(new File("presentation.pptx"), 4);
 * ImageIO.write(image, "png", new File("slide4.png"));
 *
 * // ou directement, au format choisi via RenderOptions (PNG par defaut) :
 * PptxSlideRenderer.renderSlideToFile(new File("presentation.pptx"), 4, new File("slide4.png"));
 *
 * RenderOptions svgOptions = RenderOptions.builder().format(OutputFormat.SVG).build();
 * PptxSlideRenderer.renderSlideToFile(new File("presentation.pptx"), 4, new File("slide4.svg"), svgOptions);
 * }</pre>
 *
 * <h2>Fidelite du rendu et limites connues</h2>
 * <p>Le texte, les formes vectorielles, les images et les tableaux natifs
 * sont rendus fidelement. Limites documentees :
 * <ul>
 *   <li><b>Les graphiques integres (Excel/{@code XSLFChart}) ne sont pas rendus</b>
 *       par Apache POI - un espace vide apparait a leur emplacement. Contournement
 *       possible non implemente ici : extraire les donnees via l'API {@code XSLFChart}/
 *       {@code XDDFChart} et dessiner le graphique separement (ex. avec JFreeChart).</li>
 *   <li><b>Ecart de metriques de police</b> : pour certaines polices, Java2D/AWT
 *       surestime la hauteur de texte necessaire (jusqu'a 30-35% observe dans certains
 *       cas) par rapport au moteur de rendu natif de PowerPoint, ce qui peut faire
 *       deborder une zone de texte hors de sa boite.
 *       Corrige par defaut (voir {@link RenderOptions#isFixTextOverflow()}) en
 *       retrecissant la police des formes concernees ({@code NORMAL} et
 *       {@code SHAPE} systematiquement, {@code NONE} uniquement lorsque le
 *       debordement chevaucherait reellement une autre forme de texte voisine) -
 *       desactivable via {@link RenderOptions.Builder#fixTextOverflow(boolean)}.</li>
 *   <li><b>SVG : rendu du texte dependant des polices installees chez le lecteur</b> -
 *       voir {@link OutputFormat#SVG}.</li>
 * </ul>
 */
public final class PptxSlideRenderer {

    private static final Logger LOG = LoggerFactory.getLogger(PptxSlideRenderer.class);

    /** Espace de noms XML du format SVG, requis pour creer le document DOM cible de Batik. */
    private static final String SVG_NAMESPACE_URI = "http://www.w3.org/2000/svg";

    private PptxSlideRenderer() {
    }

    /** Equivalent a {@link #renderSlide(File, int, RenderOptions)} avec {@link RenderOptions#defaults()}. */
    public static BufferedImage renderSlide(File pptxFile, int slideIndex) throws PptxRenderException {
        return renderSlide(pptxFile, slideIndex, RenderOptions.defaults());
    }

    /**
     * Rend une slide en image raster. {@link RenderOptions#getFormat()} n'a pas
     * d'effet ici (toujours un {@link BufferedImage}) : il n'est pris en compte
     * que par {@link #renderSlideToFile(File, int, File, RenderOptions)}.
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
        return renderWithSlide(pptxFile, slideIndex, options, (slide, pageSize, width, height) -> {
            BufferedImage img = new BufferedImage(width, height, BufferedImage.TYPE_INT_ARGB);
            Graphics2D graphics = img.createGraphics();
            try {
                paintSlide(graphics, slide, pageSize, options, width, height, slideIndex);
            } finally {
                graphics.dispose();
            }
            return img;
        });
    }

    /**
     * Rend une slide en SVG (rendu vectoriel, via Apache Batik) et retourne le
     * document SVG complet sous forme de chaine de caracteres. Reutilise le
     * meme pipeline de dessin que {@link #renderSlide} (y compris le correctif
     * de debordement de texte) - voir {@link OutputFormat#SVG} pour la limite
     * specifique a ce format (dependance aux polices installees chez le lecteur).
     *
     * @throws PptxRenderException si le fichier est illisible/invalide, si
     *                              {@code slideIndex} est hors bornes, ou en
     *                              cas d'erreur interne de rendu/serialisation.
     */
    public static String renderSlideAsSvg(File pptxFile, int slideIndex, RenderOptions options) throws PptxRenderException {
        return renderWithSlide(pptxFile, slideIndex, options, (slide, pageSize, width, height) -> {
            DOMImplementation domImpl = GenericDOMImplementation.getDOMImplementation();
            Document document = domImpl.createDocument(SVG_NAMESPACE_URI, "svg", null);
            SVGGraphics2D svgGraphics = new SVGGraphics2D(document);
            svgGraphics.setSVGCanvasSize(new Dimension(width, height));
            try {
                paintSlide(svgGraphics, slide, pageSize, options, width, height, slideIndex);
                try (Writer writer = new java.io.StringWriter()) {
                    // useCSS=true : styles factorises en attributs "style" plutot que
                    // repetes sur chaque element, document SVG plus compact.
                    svgGraphics.stream(writer, true);
                    return writer.toString();
                }
            } finally {
                svgGraphics.dispose();
            }
        });
    }

    /** Equivalent a {@link #renderSlideToFile(File, int, File, RenderOptions)} avec {@link RenderOptions#defaults()} (produit un PNG). */
    public static void renderSlideToFile(File pptxFile, int slideIndex, File outputFile) throws PptxRenderException {
        renderSlideToFile(pptxFile, slideIndex, outputFile, RenderOptions.defaults());
    }

    /**
     * Rend une slide et ecrit directement le resultat dans {@code outputFile},
     * au format determine par {@link RenderOptions#getFormat()} (PNG par defaut).
     *
     * @throws PptxRenderException si le rendu echoue, ou si l'ecriture du
     *                              fichier de sortie echoue.
     */
    public static void renderSlideToFile(File pptxFile, int slideIndex, File outputFile, RenderOptions options) throws PptxRenderException {
        Objects.requireNonNull(outputFile, "outputFile");
        Objects.requireNonNull(options, "options");

        switch (options.getFormat()) {
            case SVG: {
                String svg = renderSlideAsSvg(pptxFile, slideIndex, options);
                try (Writer writer = new OutputStreamWriter(new FileOutputStream(outputFile), StandardCharsets.UTF_8)) {
                    writer.write(svg);
                } catch (IOException e) {
                    throw new PptxRenderException("Impossible d'ecrire le fichier SVG de sortie: " + outputFile, e);
                }
                break;
            }
            case JPEG: {
                BufferedImage jpegImage = renderSlide(pptxFile, slideIndex, options);
                writeJpeg(jpegImage, outputFile, options.getJpegQuality());
                break;
            }
            case PNG:
            default: {
                BufferedImage pngImage = renderSlide(pptxFile, slideIndex, options);
                try (FileOutputStream out = new FileOutputStream(outputFile)) {
                    ImageIO.write(pngImage, "png", out);
                } catch (IOException e) {
                    throw new PptxRenderException("Impossible d'ecrire l'image PNG de sortie: " + outputFile, e);
                }
                break;
            }
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

    /**
     * Peint le contenu d'une slide (fond, correctif de debordement de texte,
     * dessin POI) sur un {@link Graphics2D} arbitraire - fonctionne aussi bien
     * avec un {@code Graphics2D} raster (issu d'un {@link BufferedImage}) qu'avec
     * le {@code SVGGraphics2D} vectoriel de Batik, qui implemente le meme
     * contrat. C'est ce partage qui garantit que le correctif de debordement de
     * texte s'applique identiquement, quel que soit le format de sortie.
     */
    private static void paintSlide(Graphics2D graphics, XSLFSlide slide, Dimension pageSize,
                                    RenderOptions options, int width, int height, int slideIndex) {
        graphics.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
        graphics.setRenderingHint(RenderingHints.KEY_TEXT_ANTIALIASING, RenderingHints.VALUE_TEXT_ANTIALIAS_ON);
        graphics.setRenderingHint(RenderingHints.KEY_INTERPOLATION, RenderingHints.VALUE_INTERPOLATION_BICUBIC);

        float scale = options.getScale();
        graphics.setPaint(options.getBackground());
        graphics.fill(new Rectangle2D.Float(0, 0, width, height));
        graphics.scale(scale, scale);
        graphics.fill(new Rectangle2D.Float(0, 0, pageSize.width, pageSize.height));

        int symbolRunsFixed = SymbolFontRunFixer.fixMixedSymbolRuns(slide);
        if (symbolRunsFixed > 0 && LOG.isDebugEnabled()) {
            LOG.debug("{} run(s) de police symbole (Wingdings/Webdings...) scinde(s) pour corriger du texte illisible (slide {})",
                    symbolRunsFixed, slideIndex);
        }

        if (options.isFixTextOverflow()) {
            int shrunk = OverflowAwareTextFitter.fitOverflowingText(slide, graphics);
            if (shrunk > 0 && LOG.isDebugEnabled()) {
                LOG.debug("{} forme(s) de texte retrecie(s) pour corriger un debordement (slide {})", shrunk, slideIndex);
            }
        }

        slide.draw(graphics);
    }

    /**
     * Ecrit une image raster au format JPEG. Convertit prealablement en RGB
     * opaque : le writer JPEG standard ne supporte pas le canal alpha (une
     * image {@code TYPE_INT_ARGB} le fait echouer), et le fond configure via
     * {@link RenderOptions#getBackground()} a de toute facon deja ete peint de
     * maniere opaque lors du rendu (voir {@link #paintSlide}).
     */
    private static void writeJpeg(BufferedImage image, File outputFile, float quality) throws PptxRenderException {
        BufferedImage rgbImage = new BufferedImage(image.getWidth(), image.getHeight(), BufferedImage.TYPE_INT_RGB);
        Graphics2D g = rgbImage.createGraphics();
        try {
            g.drawImage(image, 0, 0, null);
        } finally {
            g.dispose();
        }

        Iterator<ImageWriter> writers = ImageIO.getImageWritersByFormatName("jpg");
        if (!writers.hasNext()) {
            throw new PptxRenderException("Aucun writer JPEG disponible dans cet environnement Java");
        }
        ImageWriter writer = writers.next();
        try {
            ImageWriteParam param = writer.getDefaultWriteParam();
            if (param.canWriteCompressed()) {
                param.setCompressionMode(ImageWriteParam.MODE_EXPLICIT);
                param.setCompressionQuality(quality);
            }
            try (ImageOutputStream ios = ImageIO.createImageOutputStream(outputFile)) {
                writer.setOutput(ios);
                writer.write(null, new IIOImage(rgbImage, null, null), param);
            }
        } catch (IOException e) {
            throw new PptxRenderException("Impossible d'ecrire l'image JPEG de sortie: " + outputFile, e);
        } finally {
            writer.dispose();
        }
    }

    /**
     * Ouvre le fichier pptx, valide {@code slideIndex}, calcule la geometrie de
     * rendu (taille de page, dimensions physiques mises a l'echelle) puis
     * delegue le rendu proprement dit a {@code task} - factorise le
     * boilerplate d'ouverture/validation/gestion d'erreur commun a
     * {@link #renderSlide} et {@link #renderSlideAsSvg}.
     */
    private static <T> T renderWithSlide(File pptxFile, int slideIndex, RenderOptions options, SlideRenderTask<T> task)
            throws PptxRenderException {
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

            XSLFSlide slide = slides.get(slideIndex - 1);
            return task.render(slide, pageSize, width, height);
        } catch (PptxRenderException e) {
            throw e;
        } catch (IOException e) {
            throw new PptxRenderException("Impossible de lire le fichier pptx: " + pptxFile, e);
        } catch (RuntimeException e) {
            // Apache POI (parsing OOXML, rendu Graphics2D) et Batik (serialisation SVG)
            // levent parfois des exceptions non verifiees (format invalide, forme non
            // supportee...) : on les enveloppe pour garder une API d'erreur unique et previsible.
            throw new PptxRenderException("Erreur lors du rendu de la slide " + slideIndex + " de " + pptxFile, e);
        }
    }

    /** Callback de rendu utilise par {@link #renderWithSlide}, une fois le fichier ouvert et l'index valide. */
    @FunctionalInterface
    private interface SlideRenderTask<T> {
        T render(XSLFSlide slide, Dimension pageSize, int width, int height) throws IOException;
    }
}
