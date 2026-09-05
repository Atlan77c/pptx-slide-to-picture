package io.github.atlan77c.pptx2picture.internal;

import org.apache.poi.sl.draw.DrawFactory;
import org.apache.poi.sl.draw.DrawPictureShape;
import org.apache.poi.sl.draw.Drawable;
import org.apache.poi.sl.usermodel.PictureShape;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.openxmlformats.schemas.drawingml.x2006.main.CTAlphaModulateFixedEffect;
import org.openxmlformats.schemas.drawingml.x2006.main.CTBlip;
import org.openxmlformats.schemas.drawingml.x2006.main.CTBlipFillProperties;
import org.openxmlformats.schemas.presentationml.x2006.main.CTPicture;
import org.apache.xmlbeans.XmlCursor;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.xml.namespace.QName;
import java.awt.AlphaComposite;
import java.awt.Composite;
import java.awt.Graphics2D;
import java.util.List;

/**
 * Corrige un ecart de fidelite decouvert sur un fichier reel : une image
 * rendue partiellement transparente dans PowerPoint (poignee "Transparence"
 * du volet Format de l'image) ressort entierement opaque au rendu, masquant
 * tout ce qui devrait se voir en transparence derriere elle.
 *
 * <p><b>Constat sur le fichier reel</b> (slide 12 de
 * "Mes Evenements Emploi et Prestation - Doc vision 1.0.pptx") : l'image de
 * gauche porte, dans PowerPoint, une transparence de 66% (soit 34% d'opacite)
 * - confirme cote source par l'attribut XML {@code <a:blipFill><a:blip
 * r:embed="rId5"><a:alphaModFix amt="34000"/></a:blip>...</a:blipFill>} sur
 * le {@code <p:pic>} correspondant. Le PNG produit par cette librairie
 * affiche pourtant l'image totalement opaque.
 *
 * <p><b>Cause racine, identifiee en lisant directement le code source
 * d'Apache POI</b> ({@code org.apache.poi.sl.draw}, verifie sur la version
 * 5.2.5, celle utilisee par ce projet - voir {@code pom.xml}) : {@code
 * DrawPictureShape.drawContent(Graphics2D)} charge l'image puis appelle
 * directement {@code renderer.drawImage(graphics, anchor, insets)}, sans
 * jamais lire ni appliquer {@code <a:alphaModFix>}. Point notable : {@code
 * ImageRenderer.setAlpha(double)} ("the alpha [0..1] to be added to the
 * image") existe deja comme methode publique de l'interface dans cette meme
 * version 5.2.5 - c'est uniquement {@code drawContent(...)} qui ne l'appelle
 * jamais, la valeur d'{@code alphaModFix} n'etant simplement lue nulle part
 * dans le code de rendu d'une image.
 *
 * <p>Ce manque est deja corrige cote Apache POI - voir la Pull Request
 * <a href="https://github.com/apache/poi/pull/990">#990 "Support rendering
 * transparent bitmaps in presentations"</a>, qui ajoute {@code
 * PictureShape.getAlpha()} (lu depuis {@code <a:alphaModFix>}, echelle 0 a
 * 100 000 - {@code 100000} = pleinement opaque, exactement l'echelle native
 * de l'attribut XML {@code amt}) et fait appeler {@code renderer.setAlpha(...)}
 * par {@code drawContent(...)}. Mais cette PR n'a ete fusionnee sur la
 * branche de developpement d'Apache POI que le 17 janvier 2026 - posterieure
 * a la derniere version publiee a ce jour, 5.5.1 (30 novembre 2025) : aucune
 * version publiee d'Apache POI n'inclut donc ce correctif pour l'instant,
 * d'ou la necessite de le reproduire localement en attendant une future
 * version publiee qui l'integre (a ce moment-la, ce fixer pourra etre
 * retire).
 *
 * <p><b>Correctif retenu</b> : lire directement {@code amt} sur le premier
 * {@code <a:alphaModFix>} du {@code <a:blip>} de l'image (via l'API OOXML
 * publique d'Apache POI - {@link XSLFShape#getXmlObject()} suivi de la meme
 * chaine de navigation {@code CTPicture -&gt; CTBlipFillProperties -&gt;
 * CTBlip} que celle qu'Apache POI utilise lui-meme en interne ailleurs, ex.
 * pour resoudre l'identifiant de relation de l'image), puis, si {@code amt}
 * est strictement inferieur a 100 000 (image pas pleinement opaque - meme
 * condition que la future version officielle), poser un {@link
 * AlphaComposite} ({@code SRC_OVER}, fraction {@code amt / 100000f}) sur le
 * {@code Graphics2D} juste avant de peindre l'image, puis le restaurer aussitot
 * apres - plutot que d'appeler {@code ImageRenderer.setAlpha(...)} comme le
 * fera la future version officielle. Ce choix evite de devoir reimplementer
 * la boucle de chargement d'image de {@code DrawPictureShape.drawContent(...)}
 * (detection du type de fichier, image alternative de repli...) : il suffit
 * d'envelopper l'appel a {@code drawContent(...)} deja delegue - meme
 * mecanisme d'enveloppe que {@link PictureGeometryClipFixer}, qui pose de la
 * meme facon un clip sur {@code graphics} avant de deleguer. Fonctionne
 * identiquement pour le PNG/JPEG (raster) et le SVG (Apache Batik) : {@code
 * SVGGraphics2D} traduit nativement un {@link AlphaComposite} pose sur
 * {@code Graphics2D} en attribut d'opacite SVG sur le groupe englobant -
 * contrairement a {@code ImageRenderer.setAlpha(...)}, qui n'agirait que sur
 * les pixels d'un rendu raster.
 *
 * <p><b>Compose avec {@link PictureGeometryClipFixer}, jamais a sa place</b> :
 * la decoupe selon une forme et la transparence sont deux proprietes
 * independantes d'une meme image (ex. une photo a la fois recadree en
 * ellipse ET semi-transparente) - voir {@link #wrap(PictureShape,
 * DrawPictureShape)}, qui enveloppe le {@code DrawPictureShape} DEJA choisi
 * par {@link PictureGeometryClipFixer} (celui-ci, ou le comportement standard
 * de POI si l'image est rectangulaire) plutot que de se substituer a lui -
 * voir {@link DrawFactoryComposer}, seul point d'installation reel dans cette
 * librairie.
 *
 * <p><b>Detail d'implementation impose par {@code poi-ooxml-lite}</b> : cette
 * dependance transitive par defaut de {@code poi-ooxml} ne genere que les
 * classes de schema OOXML "habituellement utilisees", telles qu'identifiees
 * par les propres tests unitaires d'Apache POI (voir
 * <a href="https://poi.apache.org/help/faq.html">la FAQ officielle</a>) - un
 * type qu'Apache POI ne lit jamais lui-meme en interne (le cas de {@code
 * CTAlphaModulateFixedEffect}/{@code <a:alphaModFix>}, avant ce fixer) peut y
 * etre genere de facon incomplete. Constate en pratique : {@code
 * CTAlphaModulateFixedEffect.getAmt()} n'y renvoie pas {@code int} comme
 * attendu (erreur de compilation "incompatible types: java.lang.Object
 * cannot be converted to int"). Plutot que de forcer une dependance
 * supplementaire sur {@code poi-ooxml-full} (rejete : indisponible dans
 * certains environnements reseau restreints, alors que cette bibliotheque
 * n'a par ailleurs besoin d'aucune autre classe de {@code poi-ooxml-full}),
 * {@link #resolveAlphaModFixAmt} lit l'attribut {@code amt} directement via
 * un {@link XmlCursor} ({@code getAttributeText(QName)}) plutot que par cet
 * accesseur type - {@code XmlCursor} est herite de {@code
 * org.apache.xmlbeans.XmlObject}, fourni par la bibliotheque XmlBeans
 * elle-meme (jamais allegee, contrairement aux classes de schema OOXML
 * generees), et reste donc pleinement fonctionnel quel que soit le jar de
 * schemas utilise.
 *
 * <p><b>Limite assumee</b> : suppose qu'aucun autre correctif de ce paquetage
 * ne modifie deja le {@code Composite} du {@code Graphics2D} avant l'appel
 * (vrai a ce jour - seul {@link PictureGeometryClipFixer} touche a l'etat du
 * {@code Graphics2D} pour les images, et seulement son clip, jamais son
 * Composite) : le Composite precedent est restaure tel quel apres coup, sans
 * tenter de le combiner avec la fraction d'opacite d'{@code alphaModFix}.
 */
public final class PictureAlphaModFixer {

    private static final Logger LOG = LoggerFactory.getLogger(PictureAlphaModFixer.class);

    /**
     * Valeur de l'attribut {@code amt} d'{@code <a:alphaModFix>} correspondant
     * a une image pleinement opaque (echelle OOXML 0-100000, voir Javadoc de
     * la classe) - une image sans effet declare est implicitement a cette
     * valeur.
     */
    private static final int FULLY_OPAQUE_AMT = 100_000;

    private PictureAlphaModFixer() {
    }

    /**
     * A appeler juste avant {@code slide.draw(graphics)} - installe, seul (pas
     * en composition avec {@link ConnectorArrowFixer}/{@link
     * PictureGeometryClipFixer}), un {@link DrawFactory} qui n'applique QUE ce
     * correctif. Reserve aux tests unitaires isoles de cette classe ; le
     * rendu reel de cette librairie passe par {@link DrawFactoryComposer},
     * qui compose ce correctif avec {@link PictureGeometryClipFixer} - voir
     * Javadoc de la classe.
     *
     * @return la valeur precedente du hint {@code Drawable.DRAW_FACTORY} sur
     * ce {@code Graphics2D} (peut etre {@code null}), a repasser telle quelle
     * a {@link #restoreAfterDraw} une fois {@code slide.draw(graphics)}
     * termine.
     */
    public static Object installBeforeDraw(Graphics2D graphics) {
        Object previous = graphics.getRenderingHint(Drawable.DRAW_FACTORY);
        graphics.setRenderingHint(Drawable.DRAW_FACTORY, createDrawFactory());
        if (LOG.isDebugEnabled()) {
            LOG.debug("DrawFactory de transparence alphaModFix installe (seul, hors composition)");
        }
        return previous;
    }

    /**
     * A appeler juste apres {@code slide.draw(graphics)}, avec la valeur
     * renvoyee par {@link #installBeforeDraw} - remet {@code graphics} dans
     * l'etat ou {@link #installBeforeDraw} l'a trouve.
     */
    public static void restoreAfterDraw(Graphics2D graphics, Object previousDrawFactory) {
        graphics.setRenderingHint(Drawable.DRAW_FACTORY, previousDrawFactory);
    }

    /**
     * Cree une nouvelle instance du {@code DrawFactory} de ce correctif, sans
     * l'installer sur un {@code Graphics2D} - reserve a {@link
     * #installBeforeDraw} et aux tests.
     */
    static DrawFactory createDrawFactory() {
        return new AlphaModDrawFactory();
    }

    /**
     * Enveloppe {@code delegate} (le {@code DrawPictureShape} deja choisi par
     * ailleurs pour {@code shape} - typiquement celui de {@link
     * PictureGeometryClipFixer}, ou le comportement standard de POI) pour lui
     * appliquer, en plus, la transparence {@code alphaModFix} resolue sur
     * {@code shape} - voir Javadoc de la classe, section "Compose avec
     * PictureGeometryClipFixer". Point d'entree utilise par {@link
     * DrawFactoryComposer}.
     *
     * @return {@code delegate} tel quel si {@code shape} ne porte aucun
     * {@code alphaModFix} avec {@code amt < 100000} (aucune image deja
     * pleinement opaque n'est enveloppee inutilement), un {@code
     * DrawPictureShape} qui applique la transparence puis delegue a {@code
     * delegate} sinon.
     */
    static DrawPictureShape wrap(PictureShape<?, ?> shape, DrawPictureShape delegate) {
        Integer amt = resolveAlphaModFixAmt(shape);
        if (amt == null) {
            return delegate;
        }
        if (LOG.isDebugEnabled()) {
            LOG.debug("Image '{}' : alphaModFix amt={} ({}% d'opacite) applique via AlphaComposite "
                            + "(non gere par le rendu standard d'Apache POI 5.2.5 - voir Javadoc de la classe)",
                    shape.getShapeName(), amt, amt / 1000.0);
        }
        return new AlphaModulatedPictureShape(shape, delegate, amt);
    }

    /**
     * Resout la valeur {@code amt} du premier {@code <a:alphaModFix>} declare
     * sur le {@code <a:blip>} de {@code shape}, ou {@code null} si {@code
     * shape} n'en porte aucun, ou si l'unique/premier {@code alphaModFix}
     * declare vaut deja 100000 (pleinement opaque - rien a corriger).
     *
     * <p>Ne suit volontairement aucun heritage de placeholder (contrairement
     * a {@code PictureGeometryClipFixer.resolveGeometry}) : {@code
     * alphaModFix} est un effet applique au niveau du {@code <a:blip>}
     * lui-meme, toujours declare directement sur l'image concernee dans les
     * fichiers reels traites a ce jour - aucun cas d'heritage rencontre.
     */
    private static Integer resolveAlphaModFixAmt(PictureShape<?, ?> shape) {
        // Detour par Object avant l'instanceof/cast vers XSLFShape - meme
        // limitation du compilateur que celle documentee dans
        // PictureGeometryClipFixer.resolveGeometry (PictureShape<S,P> porte
        // des parametres generiques auto-bornes qu'un instanceof/cast direct
        // depuis le joker <?, ?> ne peut pas traverser).
        Object rawShape = shape;
        if (!(rawShape instanceof XSLFShape)) {
            // Defensif : cette bibliotheque ne traite que le format .pptx (XSLF).
            return null;
        }

        Object rawXml = ((XSLFShape) rawShape).getXmlObject();
        if (!(rawXml instanceof CTPicture)) {
            return null;
        }
        CTBlipFillProperties blipFill = ((CTPicture) rawXml).getBlipFill();
        if (blipFill == null || !blipFill.isSetBlip()) {
            return null;
        }
        CTBlip blip = blipFill.getBlip();
        List<CTAlphaModulateFixedEffect> alphaModFixes = blip.getAlphaModFixList();
        if (alphaModFixes.isEmpty()) {
            return null;
        }

        CTAlphaModulateFixedEffect alphaModFix = alphaModFixes.get(0);
        int amt = readAmtAttribute(alphaModFix);
        if (amt >= FULLY_OPAQUE_AMT) {
            // Deja pleinement opaque : aucun changement necessaire (meme
            // condition que la future version officielle d'Apache POI, PR #990).
            return null;
        }
        return Math.max(amt, 0);
    }

    /**
     * Lit l'attribut {@code amt} de {@code alphaModFix} via un {@link
     * XmlCursor} plutot que par l'accesseur type {@code getAmt()} genere par
     * XmlBeans - ce dernier est genere de facon incomplete (renvoie {@code
     * Object} au lieu de {@code int}) dans {@code poi-ooxml-lite} - voir
     * Javadoc de la classe, section "Detail d'implementation impose par
     * poi-ooxml-lite". {@code amt} n'est jamais prefixe par un espace de noms
     * dans le XML OOXML ({@code <a:alphaModFix amt="34000"/>}), d'ou le
     * {@code QName} sans espace de noms.
     *
     * @return la valeur de {@code amt}, ou {@link #FULLY_OPAQUE_AMT} si
     * l'attribut est absent ou illisible (defensif - jamais rencontre en
     * pratique sur les fichiers reels traites, {@code amt} etant toujours
     * present des lors que l'element {@code <a:alphaModFix>} existe).
     */
    private static int readAmtAttribute(CTAlphaModulateFixedEffect alphaModFix) {
        XmlCursor cursor = alphaModFix.newCursor();
        try {
            String amtText = cursor.getAttributeText(new QName("amt"));
            if (amtText == null) {
                return FULLY_OPAQUE_AMT;
            }
            try {
                return Integer.parseInt(amtText.trim());
            } catch (NumberFormatException e) {
                if (LOG.isDebugEnabled()) {
                    LOG.debug("Attribut amt='{}' illisible sur <a:alphaModFix> - traite comme pleinement opaque", amtText);
                }
                return FULLY_OPAQUE_AMT;
            }
        } finally {
            cursor.dispose();
        }
    }

    /** {@code DrawFactory} personnalise reserve a {@link #installBeforeDraw} (usage isole, voir sa Javadoc). */
    private static final class AlphaModDrawFactory extends DrawFactory {
        @Override
        public DrawPictureShape getDrawable(PictureShape<?, ?> shape) {
            return wrap(shape, super.getDrawable(shape));
        }
    }

    /**
     * Dessinateur de remplacement qui pose un {@link AlphaComposite}
     * correspondant a {@code amt} sur le {@code Graphics2D}, delegue
     * integralement le dessin a {@code delegate} (qui peut lui-meme etre le
     * {@code DrawPictureShape} standard de POI ou un {@code
     * GeometryClippedPictureShape} de {@link PictureGeometryClipFixer}), puis
     * restaure le {@code Composite} d'origine - voir Javadoc de la classe.
     *
     * <p>Etend {@code DrawPictureShape} (plutot que d'implementer {@code
     * Drawable} directement) uniquement pour respecter le type de retour de
     * {@code DrawFactory.getDrawable(PictureShape)} ; son propre {@code
     * drawContent(...)} hors de {@link #drawContent} n'est jamais invoque
     * puisque cette classe ne delegue qu'a travers {@link #drawContent}.
     */
    private static final class AlphaModulatedPictureShape extends DrawPictureShape {

        private final DrawPictureShape delegate;
        private final float alphaFraction;

        AlphaModulatedPictureShape(PictureShape<?, ?> shape, DrawPictureShape delegate, int amt) {
            super(shape);
            this.delegate = delegate;
            this.alphaFraction = amt / (float) FULLY_OPAQUE_AMT;
        }

        @Override
        public void drawContent(Graphics2D graphics) {
            Composite previousComposite = graphics.getComposite();
            graphics.setComposite(AlphaComposite.getInstance(AlphaComposite.SRC_OVER, alphaFraction));
            try {
                delegate.drawContent(graphics);
            } finally {
                graphics.setComposite(previousComposite);
            }
        }
    }
}
