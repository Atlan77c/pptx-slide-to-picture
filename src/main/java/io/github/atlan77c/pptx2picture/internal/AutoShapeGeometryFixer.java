package io.github.atlan77c.pptx2picture.internal;

import org.apache.poi.sl.draw.DrawFactory;
import org.apache.poi.sl.draw.DrawShape;
import org.apache.poi.sl.draw.DrawTextShape;
import org.apache.poi.sl.draw.Drawable;
import org.apache.poi.sl.draw.geom.Context;
import org.apache.poi.sl.draw.geom.CustomGeometry;
import org.apache.poi.sl.draw.geom.Outline;
import org.apache.poi.sl.draw.geom.PathIf;
import org.apache.poi.sl.usermodel.SimpleShape;
import org.apache.poi.sl.usermodel.TextShape;
import org.apache.poi.util.Units;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.awt.Graphics2D;
import java.awt.Shape;
import java.awt.geom.AffineTransform;
import java.awt.geom.Rectangle2D;
import java.util.ArrayList;
import java.util.Collection;
import java.util.List;

/**
 * Corrige un ecart de fidelite decouvert sur un fichier reel (2026-08-29,
 * dans la foulee de {@link PictureGeometryClipFixer}) : une forme ORDINAIRE
 * (texte, forme automatique - pas une image) dont la geometrie de
 * remplissage (ex. un cercle, un panneau a coin arrondi dessine a la main)
 * n'est declaree que dans la mise en page ({@code slideLayoutN.xml}, cas
 * courant d'un espace reserve decoratif "habille" par le theme) et jamais
 * localement sur la slide, est rendue integralement rectangulaire - un rond
 * violet devient un carre, un panneau a sommet arrondi devient un rectangle.
 *
 * <p><b>Distinct de {@link PictureGeometryClipFixer}</b>, qui corrige un tout
 * autre symptome (une IMAGE ignore completement son contour geometrique, y
 * compris quand il est declare localement - voir sa Javadoc). Ici, le
 * probleme n'est PAS que POI saute l'etape de remplissage : pour une forme
 * ordinaire, {@code getFillPaint()} renvoie bien une valeur non nulle (une
 * couleur unie, par exemple) et {@code DrawSimpleShape.draw(Graphics2D)}
 * calcule bien un contour de remplissage via {@code computeOutlines()} -
 * verifie dans le code source d'Apache POI 5.2.5 ({@code
 * org.apache.poi.sl.draw.DrawSimpleShape}). Le probleme est que {@code
 * computeOutlines()} lit la geometrie via {@code getShape().getGeometry()} -
 * un appel LOCAL A LA FORME DESSINEE, qui (comme deja documente en detail
 * dans {@link PlaceholderGeometryResolver}) ne suit jamais l'heritage de
 * placeholder PowerPoint (slide -&gt; layout -&gt; master). Quand la
 * geometrie n'est declaree que sur la mise en page, {@code getGeometry()}
 * retombe donc sur le repli implicite "rect" - {@code computeOutlines()}
 * construit alors un contour rectangulaire, et le remplissage suit cette
 * silhouette rectangulaire au lieu de la forme voulue.
 *
 * <p><b>Correctif retenu - meme architecture que {@link
 * PictureGeometryClipFixer}</b> (substitution via {@code DrawFactory}) :
 * {@link #installBeforeDraw} installe un {@code DrawFactory} personnalise qui
 * intercepte {@code getDrawable(TextShape)} (le type de forme le plus
 * general couvrant a la fois les zones de texte et les formes automatiques
 * dans la hierarchie de {@code DrawFactory} - verifie dans son code source :
 * {@code DrawAutoShape} n'existe pas comme type distinct cote dessin, {@code
 * XSLFAutoShape} est dessine via {@code DrawTextShape} comme n'importe quelle
 * autre forme a corps de texte) pour les formes dont la geometrie resolue
 * (via {@link PlaceholderGeometryResolver#resolve}, localement ou par
 * heritage de placeholder) n'est pas un simple rectangle, et leur substitue
 * {@link GeometryResolvedTextShape} - une reimplementation de {@code
 * computeOutlines(Graphics2D)} utilisant cette geometrie resolue plutot que
 * {@code getShape().getGeometry()}. Aucun clip a poser ici (contrairement a
 * {@link PictureGeometryClipFixer}) : le contour resolu remplace directement
 * celui, errone, que POI aurait calcule - {@code DrawSimpleShape.draw()}
 * remplit et trace ensuite ce contour exactement comme d'habitude.
 *
 * <p><b>Formes volontairement laissees intactes</b> : une forme dont le type
 * resolu est {@code null} ou {@code ShapeType.RECT} continue de suivre le
 * comportement standard de POI, sans le moindre risque de regression - un
 * contour rectangulaire strictement egal a l'ancre de la forme ne changerait
 * de toute facon aucun pixel par rapport au calcul standard.
 */
public final class AutoShapeGeometryFixer {

    private static final Logger LOG = LoggerFactory.getLogger(AutoShapeGeometryFixer.class);

    private AutoShapeGeometryFixer() {
    }

    /**
     * A appeler juste avant {@code slide.draw(graphics)}. Voir la Javadoc de
     * la classe pour le detail du mecanisme, et {@link DrawFactoryComposer}
     * pour la combinaison avec les autres correctifs bases sur {@code
     * DrawFactory} - {@code Drawable.DRAW_FACTORY} ne retient qu'une seule
     * valeur a la fois.
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
            LOG.debug("DrawFactory de resolution de geometrie heritee pour les formes ordinaires installe");
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
     * l'installer sur un {@code Graphics2D} - reserve a la composition avec
     * d'autres correctifs bases sur {@code DrawFactory} (voir {@link
     * DrawFactoryComposer}).
     */
    static DrawFactory createDrawFactory() {
        return new GeometryAwareDrawFactory();
    }

    /**
     * {@code DrawFactory} personnalise qui substitue {@link
     * GeometryResolvedTextShape} pour les formes qualifiees (voir {@link
     * PlaceholderGeometryResolver.ResolvedGeometry#qualifiesForClipping}), et
     * delegue au comportement standard de POI ({@code super.getDrawable(...)})
     * pour tout le reste.
     */
    private static final class GeometryAwareDrawFactory extends DrawFactory {
        @Override
        public DrawTextShape getDrawable(TextShape<?, ?> shape) {
            CustomGeometry resolvedGeometry = resolveIfQualifies(shape);
            return resolvedGeometry != null
                    ? new GeometryResolvedTextShape(shape, resolvedGeometry)
                    : super.getDrawable(shape);
        }
    }

    /**
     * Resout la geometrie de remplissage de {@code shape} par heritage de
     * placeholder et renvoie le resultat SEULEMENT s'il qualifie pour une
     * correction (voir {@link PlaceholderGeometryResolver.ResolvedGeometry#qualifiesForClipping()}) -
     * {@code null} sinon (comportement standard rectangulaire de POI a
     * garder tel quel).
     *
     * <p>Extrait de {@link GeometryAwareDrawFactory#getDrawable} (utilise par
     * {@link #installBeforeDraw}, usage standalone de ce correctif) pour etre
     * aussi reutilisable par {@link DrawFactoryComposer}, qui compose cette
     * DECISION avec {@link AutoNumberLevelFixer#wrap} plutot qu'avec le
     * {@code DrawTextShape} concret produit ici - voir la Javadoc de {@link
     * AutoNumberLevelFixer}, section "Composition avec AutoShapeGeometryFixer",
     * pour la raison (la numerotation doit se corriger sur TOUTE forme de
     * texte, y compris celles - la grande majorite - dont la geometrie
     * resolue est un simple rectangle, donc ignorees ici).
     */
    static CustomGeometry resolveIfQualifies(TextShape<?, ?> shape) {
        // Detour par Object avant l'instanceof/cast vers SimpleShape<?, ?> - voir le
        // commentaire equivalent dans PlaceholderGeometryResolver.resolve pour le
        // detail de la limitation du compilateur contournee.
        Object rawShape = shape;
        if (!(rawShape instanceof SimpleShape)) {
            // Defensif : toute XSLFTextShape (donc toute TextShape produite par ce
            // projet, qui ne traite que le format .pptx/XSLF) implemente aussi
            // SimpleShape - ce cas ne devrait jamais se produire en pratique.
            return null;
        }
        PlaceholderGeometryResolver.ResolvedGeometry resolved =
                PlaceholderGeometryResolver.resolve((SimpleShape<?, ?>) rawShape);
        if (!resolved.qualifiesForClipping()) {
            return null;
        }
        if (LOG.isDebugEnabled()) {
            String origin = (rawShape instanceof XSLFShape) && ((XSLFShape) rawShape).isPlaceholder()
                    ? "heritee d'un placeholder" : "locale";
            String typeLabel = resolved.type != null ? resolved.type.toString() : "forme libre (custGeom)";
            LOG.debug("Forme '{}' : geometrie non rectangulaire resolue ({}, {}) - contour de "
                    + "remplissage corrige", shape.getShapeName(), typeLabel, origin);
        }
        return resolved.geometry;
    }

    /**
     * Dessinateur de remplacement pour une forme dont le contour de
     * remplissage doit venir d'une geometrie resolue par heritage de
     * placeholder - sous-classe de {@code DrawTextShape} (donc de {@code
     * DrawSimpleShape}), pour ne changer QUE le contour utilise par le
     * remplissage/trace deja standard de {@code DrawSimpleShape.draw()},
     * sans toucher au dessin du texte lui-meme ({@code drawContent},
     * herite tel quel de {@code DrawTextShape}).
     */
    private static final class GeometryResolvedTextShape extends DrawTextShape {

        private final CustomGeometry resolvedGeometry;

        GeometryResolvedTextShape(TextShape<?, ?> shape, CustomGeometry resolvedGeometry) {
            super(shape);
            this.resolvedGeometry = resolvedGeometry;
        }

        /**
         * Reimplementation de {@code DrawSimpleShape.computeOutlines
         * (Graphics2D)} (POI 5.2.5) utilisant {@link #resolvedGeometry} - deja
         * resolue par heritage de placeholder si necessaire - a la place de
         * {@code getShape().getGeometry()}. L'ancre utilisee reste celle,
         * locale, de la forme elle-meme : seule la geometrie (le contour) peut
         * venir d'un placeholder herite, jamais la position/taille.
         */
        @Override
        protected Collection<Outline> computeOutlines(Graphics2D graphics) {
            return computeResolvedOutlines(graphics, getShape(), resolvedGeometry);
        }
    }

    /**
     * Corps de {@link GeometryResolvedTextShape#computeOutlines} extrait en
     * methode statique reutilisable, pour que {@link AutoNumberLevelFixer}
     * puisse produire le meme contour resolu sans dupliquer ce calcul - voir
     * Javadoc de {@link AutoNumberLevelFixer}, section "Composition avec
     * AutoShapeGeometryFixer". Comportement strictement identique a
     * l'implementation d'origine (extraction pure, aucun changement).
     */
    static Collection<Outline> computeResolvedOutlines(Graphics2D graphics, SimpleShape<?, ?> ss,
                                                         CustomGeometry resolvedGeometry) {
        List<Outline> lst = new ArrayList<>();
        if (resolvedGeometry == null) {
            return lst;
        }

        Rectangle2D anchor = DrawShape.getAnchor(graphics, ss);
        if (anchor == null) {
            return lst;
        }

        for (PathIf p : resolvedGeometry) {
            double w = p.getW(), h = p.getH(), scaleX, scaleY;
            if (w == -1) {
                w = Units.toEMU(anchor.getWidth());
                scaleX = Units.toPoints(1);
            } else if (anchor.getWidth() == 0) {
                scaleX = 1;
            } else {
                scaleX = anchor.getWidth() / w;
            }
            if (h == -1) {
                h = Units.toEMU(anchor.getHeight());
                scaleY = Units.toPoints(1);
            } else if (anchor.getHeight() == 0) {
                scaleY = 1;
            } else {
                scaleY = anchor.getHeight() / h;
            }

            // les guides de la geometrie sont tous definis les uns
            // relativement aux autres, on construit donc le chemin a
            // partir de (0,0), exactement comme DrawSimpleShape.computeOutlines.
            final Rectangle2D pathAnchor = new Rectangle2D.Double(0, 0, w, h);

            Context ctx = new Context(resolvedGeometry, pathAnchor, ss);
            Shape gp = p.getPath(ctx);

            AffineTransform at = new AffineTransform();
            at.translate(anchor.getX(), anchor.getY());
            at.scale(scaleX, scaleY);

            Shape canvasShape = at.createTransformedShape(gp);

            lst.add(new Outline(canvasShape, p));
        }

        return lst;
    }
}
