package io.github.atlan77c.pptx2picture;

import io.github.atlan77c.pptx2picture.internal.BulletSymbolFontFixer;
import io.github.atlan77c.pptx2picture.internal.DrawFactoryComposer;
import io.github.atlan77c.pptx2picture.internal.NeighborShapeOverlapFixer;
import io.github.atlan77c.pptx2picture.internal.OverflowAwareTextFitter;
import io.github.atlan77c.pptx2picture.internal.OversizedWhitespaceRunFixer;
import io.github.atlan77c.pptx2picture.internal.RightAlignedIndentFixer;
import io.github.atlan77c.pptx2picture.internal.RoundedShapeAnchorFixer;
import io.github.atlan77c.pptx2picture.internal.SymbolFontRunFixer;
import io.github.atlan77c.pptx2picture.internal.TableCellLineSpacingFixer;
import io.github.atlan77c.pptx2picture.internal.TitleRepainter;
import org.apache.batik.dom.GenericDOMImplementation;
import org.apache.batik.svggen.SVGGraphics2D;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
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
 *       deborder une zone de texte hors de sa boite - y compris une cellule de tableau
 *       natif hors de sa propre ligne.
 *       Corrige par defaut (voir {@link RenderOptions#isFixTextOverflow()}) en
 *       retrecissant la police des formes concernees ({@code NORMAL} et
 *       {@code SHAPE} systematiquement, {@code NONE} uniquement lorsque le
 *       debordement chevaucherait reellement une autre forme de texte voisine ;
 *       une cellule de tableau, en corrigeant d'abord son interligne) -
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

        // [Diagnostic 2026-09-05, a retirer une fois le diagnostic termine] voir Javadoc de
        // logTextSnapshot() - encadre tous les correctifs "avant dessin" pour determiner si un
        // texte manquant au rendu final est deja absent du modele a ce stade (regression d'un
        // correctif de ce paquetage) ou seulement absent du dessin produit par slide.draw()
        // (cause hors de ce paquetage). Voir conversion_pptx_vers_images.md, section 26.
        logTextSnapshot(slide, graphics, "initial (avant tout correctif)", slideIndex);

        int symbolRunsFixed = SymbolFontRunFixer.fixMixedSymbolRuns(slide);
        if (symbolRunsFixed > 0 && LOG.isDebugEnabled()) {
            LOG.debug("{} run(s) de police symbole (Wingdings/Webdings...) scinde(s) pour corriger du texte illisible (slide {})",
                    symbolRunsFixed, slideIndex);
        }

        // Independant de SymbolFontRunFixer ci-dessus : celui-ci corrige les polices
        // symbole posees sur des RUNS de texte (<a:sym>), pas les puces de paragraphe
        // (<a:buFont>/<a:buChar>), un chemin de rendu POI distinct - voir Javadoc de
        // BulletSymbolFontFixer.
        int bulletFontsFixed = BulletSymbolFontFixer.fixSymbolBulletFonts(slide);
        if (bulletFontsFixed > 0 && LOG.isDebugEnabled()) {
            LOG.debug("{} puce(s) avec police symbole non reconnue (liste de repli type \"Wingdings,Sans-Serif\") "
                            + "normalisee(s) pour corriger un pictogramme denature (slide {})",
                    bulletFontsFixed, slideIndex);
        }

        // Avant OverflowAwareTextFitter : un run blanc surdimensionne gonfle la hauteur de
        // ligne mesuree par Java2D pour un interligne en pourcentage, faussant toute mesure
        // de hauteur faite ensuite (y compris par OverflowAwareTextFitter). Voir Javadoc de
        // OversizedWhitespaceRunFixer.
        int whitespaceRunsFixed = OversizedWhitespaceRunFixer.fixOversizedWhitespaceRuns(slide);
        if (whitespaceRunsFixed > 0 && LOG.isDebugEnabled()) {
            LOG.debug("{} run(s) blanc(s) surdimensionne(s) (taille de police superieure au texte visible du "
                            + "meme paragraphe) ramene(s) a une taille normale (slide {})",
                    whitespaceRunsFixed, slideIndex);
        }

        // Avant OverflowAwareTextFitter : ce dernier lit ts.getVerticalAlignment() pour
        // calculer ses zones de debordement, donc l'ancrage doit deja etre corrige a ce stade.
        int anchorsFixed = RoundedShapeAnchorFixer.fixVerticalAnchor(slide);
        if (anchorsFixed > 0 && LOG.isDebugEnabled()) {
            LOG.debug("{} forme(s) spAutoFit non rectangulaire(s) recentree(s) verticalement (slide {})",
                    anchorsFixed, slideIndex);
        }

        // Independant de RoundedShapeAnchorFixer/OverflowAwareTextFitter (touche marL/indent,
        // pas l'ancrage vertical ni la taille de police) - avant slide.draw(graphics) comme les
        // deux precedents. Voir Javadoc de RightAlignedIndentFixer.
        int indentsFixed = RightAlignedIndentFixer.fixInheritedIndent(slide);
        if (indentsFixed > 0 && LOG.isDebugEnabled()) {
            LOG.debug("{} paragraphe(s) aligne(s) a droite/centre avec un retrait herite du masque corrige(s) (slide {})",
                    indentsFixed, slideIndex);
        }

        if (options.isFixTextOverflow()) {
            // options.isBroadenAutofitExemption() : elargissement EXPERIMENTAL (2026-09-05,
            // desactive par defaut) de l'exemption d'autofit mal classe par POI - voir Javadoc
            // de OverflowAwareTextFitter, section "Elargissement general (experimental)", et
            // section 26 du markdown de suivi (slide 16, "fichier-test-B.pptx"). Passe identiquement aux deux correctifs ci-dessous (meme
            // exemption partagee, voir OverflowAwareTextFitter#isAutofitBroadeningExempt).
            boolean broadenAutofitExemption = options.isBroadenAutofitExemption();
            int shrunk = OverflowAwareTextFitter.fitOverflowingText(slide, graphics, broadenAutofitExemption);
            if (shrunk > 0 && LOG.isDebugEnabled()) {
                LOG.debug("{} forme(s) de texte retrecie(s) pour corriger un debordement (slide {})", shrunk, slideIndex);
            }

            // Complementaire, jamais redondant, avec OverflowAwareTextFitter : ce
            // dernier ne traite que le debordement d'une forme au-dela de SA PROPRE
            // ancre. NeighborShapeOverlapFixer couvre l'angle mort restant - un texte
            // qui tient dans sa propre ancre (volontairement surdimensionnee pour
            // accueillir une forme voisine posee par-dessus) mais dont le contenu
            // reellement mesure s'etend neanmoins assez bas pour chevaucher cette
            // forme voisine independante - y compris quand ce chevauchement ne
            // concerne, a l'echelle du bloc entier, que des paragraphes ENTIEREMENT
            // VIDES places en espaceurs devant une annotation (detection paragraphe
            // par paragraphe, pas seulement forme entiere - voir Javadoc de la classe,
            // "2e variante"). Doit s'executer apres (il ignore toute forme deja hors
            // de son perimetre, cf. Javadoc de la classe). Gouverne par le meme
            // indicateur RenderOptions.fixTextOverflow : meme famille de correction
            // (chevauchement du a une surestimation Java2D/AWT).
            int neighborFixed = NeighborShapeOverlapFixer.fixNeighborOverlaps(slide, graphics, broadenAutofitExemption);
            if (neighborFixed > 0 && LOG.isDebugEnabled()) {
                LOG.debug("{} forme(s) de texte corrigee(s) (interligne et/ou police) pour un chevauchement avec "
                        + "une forme voisine independante, sans depasser leur propre ancre (slide {})",
                        neighborFixed, slideIndex);
            }

            // Complementaire aux deux precedents, sur un perimetre totalement disjoint :
            // aucun des correctifs ci-dessus ne parcourt les cellules de tableau (voir
            // Javadoc de OverflowAwareTextFitter, section "Hors de portee : les cellules
            // de tableau natif"). Meme famille de correction (surestimation Java2D/AWT de
            // l'interligne), gouvernee par le meme indicateur RenderOptions.fixTextOverflow.
            int tableCellsFixed = TableCellLineSpacingFixer.fixTableCellLineSpacing(slide, graphics);
            if (tableCellsFixed > 0 && LOG.isDebugEnabled()) {
                LOG.debug("{} cellule(s) de tableau corrigee(s) (interligne et/ou police) pour un debordement "
                        + "sur la ligne suivante (slide {})", tableCellsFixed, slideIndex);
            }
        }

        // [v13-drawfactory-zorder][v17-picture-clip][v20-picture-alpha] Installe, via
        // DrawFactoryComposer, un DrawFactory personnalise qui substitue, pendant ce seul appel
        // a slide.draw(graphics), notre propre dessinateur pour deux categories de formes :
        //  - les connecteurs courbes/coudes a pointe de fleche declaree (POI oriente mal leur
        //    pointe - voir Javadoc de ConnectorArrowFixer) - POI continue de les dessiner a
        //    leur place naturelle dans l'ordre d'empilement des formes (contrairement aux
        //    anciennes versions, qui retiraient le connecteur de la slide pour le redessiner
        //    apres coup, ce qui le faisait passer systematiquement au premier plan quelle que
        //    soit sa position d'origine dans l'empilement) ;
        //  - les images, sur deux aspects independants et cumulables (voir Javadoc de
        //    DrawFactoryComposer, section "Cas particulier des images") : celles "decoupees
        //    selon une forme" dans PowerPoint (ex. une ellipse, une forme libre), que POI peint
        //    integralement rectangulaires, recouvrant tout ce qui se trouve derriere elles dans
        //    la boite englobante de leur ancre - voir Javadoc de PictureGeometryClipFixer - et
        //    celles rendues partiellement transparentes dans PowerPoint (poignee
        //    "Transparence" du volet Format de l'image, <a:alphaModFix>), que POI peint
        //    entierement opaques - voir Javadoc de PictureAlphaModFixer ;
        //  - les formes ORDINAIRES (texte, formes automatiques - pas des images) dont la
        //    geometrie de remplissage n'est declaree que dans la mise en page (espace reserve
        //    decoratif habille par le theme, jamais repete localement sur la slide) : POI la
        //    remplit alors integralement rectangulaire (un rond devient un carre, un panneau a
        //    coin arrondi devient un rectangle) - voir Javadoc de AutoShapeGeometryFixer.
        // DrawFactoryComposer est necessaire ici (plutot que d'appeler les installBeforeDraw de
        // chaque correctif l'un apres l'autre) car Drawable.DRAW_FACTORY est un hint a valeur
        // UNIQUE : installer un correctif independamment ecraserait silencieusement les
        // precedents - voir Javadoc de DrawFactoryComposer.
        // [Diagnostic 2026-09-05, a retirer une fois le diagnostic termine] meme methode
        // qu'au tout debut de paintSlide() - voir le commentaire a cet appel.
        logTextSnapshot(slide, graphics, "juste avant slide.draw()", slideIndex);

        Object previousDrawFactory = DrawFactoryComposer.installBeforeDraw(graphics);
        try {
            slide.draw(graphics);
        } finally {
            DrawFactoryComposer.restoreAfterDraw(graphics, previousDrawFactory);
        }

        // Apres slide.draw(graphics), a l'inverse des correctifs precedents : il ne
        // s'agit pas ici de corriger l'etat d'une forme avant le dessin, mais de
        // repeindre certaines formes PAR-DESSUS un dessin deja termine (voir Javadoc
        // de TitleRepainter).
        int titlesRepainted = TitleRepainter.repaintTitles(slide, graphics);
        if (titlesRepainted > 0 && LOG.isDebugEnabled()) {
            LOG.debug("{} titre(s) repeint(s) au premier plan (slide {})", titlesRepainted, slideIndex);
        }
    }

    /**
     * [Diagnostic 2026-09-05, a retirer une fois le diagnostic termine] Journalise, pour
     * chaque forme de texte de {@code slide}, le texte brut de chaque paragraphe (tous runs
     * concatenes) ainsi que la hauteur totale mesuree par POI ({@link
     * XSLFTextShape#getTextHeight(Graphics2D)}) comparee a la hauteur de son ancre.
     *
     * <p>But : determiner, pour un texte constate absent du rendu final (ex. une ligne de
     * fin de paragraphe qui deborde a la ligne suivante, jamais dessinee), si ce texte est
     * deja absent du MODELE POI a un moment donne (regression d'un correctif de ce
     * paquetage, si l'appel se situe apres son execution) ou seulement absent du dessin
     * produit par {@code slide.draw()} (bug hors de portee de ce paquetage, dans le
     * decoupage de ligne interne d'Apache POI) - voir {@code
     * conversion_pptx_vers_images.md}, section 26, pour le cas reel ayant motive cet ajout.
     * Un appel encadrant chaque correctif "avant dessin" permet, par simple comparaison des
     * lignes de log successives, de localiser exactement l'etape ou un texte disparaitrait
     * du modele - alors qu'aujourd'hui aucun correctif de ce paquetage ne journalise que le
     * texte AVANT/APRES son propre passage, seulement qu'il a agi ou non sur une forme.
     *
     * <p>Ne modifie jamais la forme ({@code getTextHeight(Graphics2D)} est une methode de
     * lecture pure, deja utilisee en lecture seule ailleurs dans ce paquetage - voir Javadoc
     * de {@link OverflowAwareTextFitter}) : peut etre appelee autant de fois que necessaire
     * sans aucun risque d'effet de bord sur le rendu.
     *
     * @param moment     libelle court identifiant l'etape du pipeline a laquelle cet appel a
     *                   lieu (ex. "initial (avant tout correctif)", "juste avant
     *                   slide.draw()") - affiche tel quel dans chaque ligne de log pour
     *                   distinguer les differents appels d'une meme execution.
     * @param slideIndex uniquement pour le contexte affiche dans le log (base 1, voir
     *                   {@link #renderSlide(File, int, RenderOptions)}).
     */
    private static void logTextSnapshot(XSLFSlide slide, Graphics2D graphics, String moment, int slideIndex) {
        if (!LOG.isDebugEnabled()) {
            return;
        }
        for (XSLFShape shape : slide.getShapes()) {
            if (!(shape instanceof XSLFTextShape)) {
                continue;
            }
            XSLFTextShape ts = (XSLFTextShape) shape;
            List<XSLFTextParagraph> paragraphs = ts.getTextParagraphs();
            double measuredHeight;
            try {
                measuredHeight = ts.getTextHeight(graphics);
            } catch (RuntimeException e) {
                // Mesure best-effort : ne doit jamais faire echouer le rendu reel a cause de
                // ce diagnostic. Voir Javadoc de la methode.
                measuredHeight = Double.NaN;
            }
            double anchorHeight = ts.getAnchor() != null ? ts.getAnchor().getHeight() : Double.NaN;
            LOG.debug("[diag {} | slide {}] {} : {} paragraphe(s), hauteur mesuree {}pt, ancre {}pt, wordWrap={}",
                    moment, slideIndex, shape.getShapeName(), paragraphs.size(), measuredHeight, anchorHeight,
                    ts.getWordWrap());
            for (int i = 0; i < paragraphs.size(); i++) {
                XSLFTextParagraph para = paragraphs.get(i);
                StringBuilder text = new StringBuilder();
                for (XSLFTextRun run : para.getTextRuns()) {
                    String raw = run.getRawText();
                    text.append(raw == null ? "" : raw);
                }
                LOG.debug("[diag {} | slide {}] {} - paragraphe {} (niveau {}) : \"{}\"",
                        moment, slideIndex, shape.getShapeName(), i, para.getIndentLevel(), text);
                // Diagnostic taille de police reelle par run : verifie si la taille resolue
                // (heritage layout/master inclus) correspond bien a ce qui est attendu (ex. 1600
                // pour le corps herite du master), ou si elle "explose" (ex. 1800/2000/2400) a
                // cause d'un bug de resolution d'heritage - voir section 26 du markdown de suivi.
                List<XSLFTextRun> runs = para.getTextRuns();
                if (runs.isEmpty()) {
                    LOG.debug("[diag {} | slide {}] {} - paragraphe {} : AUCUN run (ligne vide/endParaRPr)"
                                    + " - la hauteur de cette ligne depend uniquement de la taille de police par"
                                    + " defaut resolue par POI pour ce niveau/placeholder.",
                            moment, slideIndex, shape.getShapeName(), i);
                } else {
                    for (int r = 0; r < runs.size(); r++) {
                        XSLFTextRun run = runs.get(r);
                        LOG.debug("[diag {} | slide {}] {} - paragraphe {} / run {} : taille resolue={}pt,"
                                        + " police={}, gras={}, texte=\"{}\"",
                                moment, slideIndex, shape.getShapeName(), i, r, run.getFontSize(),
                                run.getFontFamily(), run.isBold(), run.getRawText());
                    }
                }
            }
        }
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
