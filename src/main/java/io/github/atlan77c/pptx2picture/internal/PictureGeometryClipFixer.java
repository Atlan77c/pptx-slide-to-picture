package io.github.atlan77c.pptx2picture.internal;

import org.apache.poi.sl.draw.DrawFactory;
import org.apache.poi.sl.draw.DrawPictureShape;
import org.apache.poi.sl.draw.Drawable;
import org.apache.poi.sl.draw.geom.Context;
import org.apache.poi.sl.draw.geom.CustomGeometry;
import org.apache.poi.sl.draw.geom.Outline;
import org.apache.poi.sl.draw.geom.PathIf;
import org.apache.poi.sl.usermodel.PictureShape;
import org.apache.poi.sl.usermodel.ShapeType;
import org.apache.poi.sl.usermodel.SimpleShape;
import org.apache.poi.util.Units;
import org.apache.poi.xslf.model.PropertyFetcher;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.awt.Graphics2D;
import java.awt.Shape;
import java.awt.geom.AffineTransform;
import java.awt.geom.Area;
import java.awt.geom.Rectangle2D;
import java.util.ArrayList;
import java.util.Collection;
import java.util.List;

/**
 * Corrige un ecart de fidelite decouvert sur un fichier reel : une image
 * inseree normalement dans PowerPoint ("Insertion &gt; Images"), puis
 * "decoupee" selon une forme via "Format de l'image &gt; Rogner &gt; Rogner
 * selon la forme" (ex. une ellipse, ou une forme libre comme une loupe), est
 * rendue integralement rectangulaire - sa geometrie visible deborde alors
 * largement du contour de la forme choisie, recouvrant tout ce qui se trouve
 * derriere elle dans la boite englobante de son ancre.
 *
 * <p><b>Cause racine, identifiee dans le code source d'Apache POI lui-meme</b>
 * ({@code org.apache.poi.sl.draw}, verifie directement sur la version 5.2.5,
 * celle utilisee par ce projet) : le dessin d'une forme se fait en deux temps
 * bien separes dans {@code DrawSimpleShape.draw(Graphics2D)} - d'abord le
 * <i>remplissage</i> de l'interieur du contour geometrique de la forme
 * ({@code computeOutlines()}, utilise uniquement si {@code getFillPaint()}
 * renvoie une valeur non nulle), puis, dans tous les cas, {@code drawContent
 * (graphics)} pour peindre le contenu propre de la forme (texte, image...).
 * Pour une image ({@code p:pic}, {@code DrawPictureShape}), {@code
 * getFillPaint(Graphics2D)} est explicitement surcharge pour renvoyer {@code
 * null} - logique, une image "brute" n'a pas de remplissage au sens de
 * DrawingML - ce qui saute purement et simplement l'etape de remplissage/
 * decoupe geometrique. {@code drawContent(Graphics2D)} dessine alors le
 * bitmap directement via {@code ImageRenderer.drawImage(graphics, anchor,
 * insets)}, ou {@code anchor} est la boite englobante RECTANGULAIRE de
 * l'image ({@code insets} ne gere que le rognage rectangulaire classique
 * {@code srcRect}, pas la forme de decoupe) : le contour geometrique de la
 * forme ({@code prstGeom}/{@code custGeom} - ellipse, arrondi, forme libre...)
 * calcule par ailleurs par {@code computeOutlines()} n'est <b>jamais appele
 * pour le cas d'une image</b>, ni utilise pour poser un clip sur {@code
 * graphics} avant {@code drawContent}. Bug confirme distinct de celui
 * corrige par {@link RoundedShapeAnchorFixer} (qui porte sur l'ancrage
 * vertical du TEXTE dans une forme non rectangulaire, jamais sur une image).
 *
 * <p><b>Correctif retenu - meme architecture que {@link ConnectorArrowFixer}
 * (substitution via {@code DrawFactory}, voir sa Javadoc pour le detail du
 * mecanisme)</b> : {@link #installBeforeDraw} installe un {@code DrawFactory}
 * personnalise qui intercepte uniquement {@code getDrawable(PictureShape)}
 * pour les images dont la geometrie declaree n'est pas un simple rectangle
 * (voir {@link #resolveGeometry}), et leur substitue {@link
 * GeometryClippedPictureShape} - qui pose, juste avant de deleguer a {@code
 * super.drawContent(graphics)}, un clip {@link Area} egal a l'union des
 * sous-chemins REMPLIS ({@code PathIf#isFilled()}) du contour de la forme -
 * seuls ceux-ci delimitent l'interieur visuel de la forme ; un sous-chemin
 * non rempli (ex. une ligne de construction interne a une forme libre) ne
 * doit pas participer a la decoupe. Le clip precedent de {@code graphics}
 * (s'il y en a un) est preserve par intersection, puis integralement restaure
 * apres coup - jamais ecrase.
 *
 * <p><b>Correction du 2026-08-26 - heritage de placeholder non resolu</b> :
 * une premiere version de ce correctif s'appuyait directement sur {@code
 * PictureShape.getShapeType()}/{@code getGeometry()} - qui, verifie sur le
 * code source d'{@code XSLFSimpleShape} (POI 5.2.5), n'inspectent QUE le
 * {@code <p:spPr>} propre a la forme dans le XML de la slide, sans jamais
 * suivre l'heritage de placeholder PowerPoint (slide -&gt; layout -&gt;
 * master). Cela n'affectait pas les tests synthetiques de ce correctif (image
 * autonome avec {@code setShapeType(ELLIPSE)} pose directement sur la slide),
 * mais s'est revele inoperant sur un cas reel tres courant : une image
 * inseree dans un <b>espace reserve image</b> ("picture placeholder", {@code
 * <p:ph type="pic" idx="N"/>}) dont la forme de decoupe (ex. l'ellipse d'un
 * theme "loupe") est definie une seule fois dans la mise en page ({@code
 * slideLayoutN.xml}, sur la forme {@code <p:sp>} idx="N" correspondante) et
 * jamais repetee localement dans le {@code <p:pic>} de la slide - PowerPoint
 * la resout par heritage a l'affichage, POI ne la resout jamais pour {@code
 * getShapeType()}/{@code getGeometry()}. Consequence observee : le correctif
 * ne se declenchait pas du tout pour ce cas (silencieusement, {@code
 * getShapeType()} renvoyant {@code null}), sans aucune erreur - exactement le
 * symptome "le correctif n'a rien change" remonte sur un fichier reel.
 * {@link #resolveGeometry} corrige cela en reproduisant la meme chaine
 * d'heritage que POI utilise deja en interne pour les autres proprietes
 * (couleur de remplissage, de trait...) - {@code XSLFShape.
 * fetchShapeProperty(PropertyFetcher)}, {@code public}, qui visite dans
 * l'ordre la forme de la slide, puis (si elle est un placeholder sans
 * geometrie locale) la forme correspondante de la mise en page, puis du
 * masque - et retient la geometrie de la premiere forme de la chaine qui en
 * declare une localement.
 *
 * <p><b>Formes volontairement laissees intactes</b> (voir {@link
 * #resolveGeometry}) : une image dont le type de forme resolu (localement ou
 * par heritage) est {@code null} ou {@link ShapeType#RECT} (le cas de la tres
 * grande majorite des images - insertion normale, ou rognage rectangulaire
 * classique {@code srcRect} sans changement de forme) continue de suivre
 * exactement le comportement standard de POI, sans le moindre risque de
 * regression : la decoupe par un rectangle strictement egal a l'ancre de
 * l'image ne changerait de toute facon aucun pixel par rapport au rendu non
 * decoupe.
 *
 * <p><b>Limite assumee</b> : {@code insets} ({@code srcRect}, le rognage
 * rectangulaire classique applicable en plus d'une decoupe par forme) reste
 * gere tel quel par {@code super.drawContent(graphics)}, seul le clip
 * geometrique supplementaire est ajoute - les deux se combinent naturellement
 * puisque le clip s'applique en dernier lieu au resultat deja rogne par
 * {@code insets}.
 */
public final class PictureGeometryClipFixer {

    private static final Logger LOG = LoggerFactory.getLogger(PictureGeometryClipFixer.class);

    private PictureGeometryClipFixer() {
    }

    /**
     * A appeler juste avant {@code slide.draw(graphics)}. Installe un {@link
     * DrawFactory} personnalise sur {@code graphics}, via le hint de rendu
     * {@code Drawable.DRAW_FACTORY}, afin que les images a geometrie non
     * rectangulaire soient decoupees selon leur forme - voir la Javadoc de la
     * classe pour le detail du mecanisme.
     *
     * <p><b>Attention</b> : ce hint ne retient qu'un seul {@code DrawFactory}
     * a la fois - pour combiner ce correctif avec {@link ConnectorArrowFixer}
     * (egalement base sur {@code DrawFactory}), utiliser {@link
     * DrawFactoryComposer#installBeforeDraw} plutot que d'appeler les deux
     * {@code installBeforeDraw} l'un apres l'autre (le second ecraserait
     * silencieusement le premier).
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
            LOG.debug("DrawFactory de decoupe des images a geometrie non rectangulaire installe");
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
        return new ClippingDrawFactory();
    }

    /**
     * Resout la geometrie et le type de forme reellement applicables a
     * {@code shape}, en suivant si besoin l'heritage de placeholder
     * PowerPoint (slide -&gt; layout -&gt; master) quand la forme elle-meme
     * ne declare localement ni {@code <a:prstGeom>} ni {@code <a:custGeom>} -
     * voir la section "Correction du 2026-08-26" de la Javadoc de la classe.
     */
    private static ResolvedGeometry resolveGeometry(PictureShape<?, ?> shape) {
        // Detour par Object avant l'instanceof/cast vers XSLFShape : PictureShape<S,P>
        // porte des parametres generiques "auto-bornes" (S extends Shape<S,P>, ...) - une
        // fois captures depuis le joker <?, ?> de la signature de getDrawable(...), javac
        // refuse un instanceof/cast direct entre ce type capture et la classe concrete
        // XSLFShape ("incompatible types : cannot be converted", constate a la compilation
        // reelle malgre cette meme conversion, sans probleme, pour XSLFShape -> SimpleShape
        // plus bas). Object n'ayant aucun parametre generique, ce detour contourne cette
        // limitation du compilateur sans changer le comportement a l'execution.
        Object rawShape = shape;
        if (!(rawShape instanceof XSLFShape)) {
            // Defensif : cette bibliotheque ne traite que le format .pptx
            // (XSLF) - ce cas ne devrait jamais se produire en pratique.
            return new ResolvedGeometry(shape.getShapeType(), shape.getGeometry());
        }
        XSLFShape xslfShape = (XSLFShape) rawShape;

        PropertyFetcher<ResolvedGeometry> fetcher = new PropertyFetcher<ResolvedGeometry>() {
            @Override
            public boolean fetch(XSLFShape candidate) {
                if (!(candidate instanceof SimpleShape)) {
                    return false;
                }
                SimpleShape<?, ?> simpleShape = (SimpleShape<?, ?>) candidate;
                ShapeType type = simpleShape.getShapeType();
                if (type == null) {
                    // Pas de <a:prstGeom>/<a:custGeom> local sur ce maillon de
                    // la chaine - fetchShapeProperty passe au suivant
                    // (layout, puis master).
                    return false;
                }
                setValue(new ResolvedGeometry(type, simpleShape.getGeometry()));
                return true;
            }
        };
        xslfShape.fetchShapeProperty(fetcher);
        return fetcher.isSet() ? fetcher.getValue() : new ResolvedGeometry(null, null);
    }

    /**
     * Resultat de {@link #resolveGeometry} : le type de forme et la
     * geometrie resolus, provenant du meme maillon de la chaine d'heritage
     * (jamais l'un de la slide et l'autre du layout).
     */
    private static final class ResolvedGeometry {
        final ShapeType type;
        final CustomGeometry geometry;

        ResolvedGeometry(ShapeType type, CustomGeometry geometry) {
            this.type = type;
            this.geometry = geometry;
        }

        /**
         * {@code true} si cette resolution designe une geometrie non
         * rectangulaire qu'il faut decouper (voir Javadoc de la classe,
         * section "Formes volontairement laissees intactes").
         */
        boolean qualifiesForClipping() {
            if (type == null || type == ShapeType.RECT) {
                return false;
            }
            return geometry != null;
        }
    }

    /**
     * {@code DrawFactory} personnalise qui substitue {@link
     * GeometryClippedPictureShape} pour les images qualifiees (voir {@link
     * ResolvedGeometry#qualifiesForClipping}), et delegue au comportement
     * standard de POI ({@code super.getDrawable(...)}) pour tout le reste.
     */
    private static final class ClippingDrawFactory extends DrawFactory {
        @Override
        public DrawPictureShape getDrawable(PictureShape<?, ?> shape) {
            ResolvedGeometry resolved = resolveGeometry(shape);
            if (resolved.qualifiesForClipping()) {
                if (LOG.isDebugEnabled()) {
                    // Detour par Object avant l'instanceof/cast vers XSLFShape - voir le
                    // commentaire de resolveGeometry ci-dessus pour le detail.
                    Object rawShape = shape;
                    String origin = (rawShape instanceof XSLFShape) && ((XSLFShape) rawShape).isPlaceholder()
                            ? "heritee d'un placeholder" : "locale";
                    LOG.debug("Image '{}' : geometrie non rectangulaire resolue ({}, {}) - decoupe appliquee",
                            shape.getShapeName(), resolved.type, origin);
                }
                return new GeometryClippedPictureShape(shape, resolved.geometry);
            }
            return super.getDrawable(shape);
        }
    }

    /**
     * Dessinateur de remplacement pour une image a geometrie non
     * rectangulaire - sous-classe de {@code DrawPictureShape} (donc de {@code
     * DrawSimpleShape}), seul moyen d'acceder a {@code getAnchor(Graphics2D,
     * PlaceableShape)}, {@code public static} chez POI mais commodement
     * herite ici.
     *
     * <p>Ne reutilise volontairement PAS {@code DrawSimpleShape.
     * computeOutlines(Graphics2D)} (qui relit {@code getShape().
     * getGeometry()} directement sur la forme dessinee, sans heritage de
     * placeholder - voir {@link #resolveGeometry}) : {@link
     * #computeResolvedOutlines} en est une reimplementation qui utilise a la
     * place la geometrie deja resolue par {@link #resolveGeometry}, transmise
     * au constructeur.
     */
    private static final class GeometryClippedPictureShape extends DrawPictureShape {

        private final CustomGeometry resolvedGeometry;

        GeometryClippedPictureShape(PictureShape<?, ?> shape, CustomGeometry resolvedGeometry) {
            super(shape);
            this.resolvedGeometry = resolvedGeometry;
        }

        @Override
        public void drawContent(Graphics2D graphics) {
            Area clipArea = buildClipArea(graphics);
            if (clipArea == null || clipArea.isEmpty()) {
                // Geometrie sans aucun sous-chemin rempli exploitable (cas degenere,
                // jamais rencontre en pratique) : par prudence, comportement standard
                // de POI plutot que de masquer entierement l'image.
                super.drawContent(graphics);
                return;
            }

            Shape previousClip = graphics.getClip();
            Area newClip = clipArea;
            if (previousClip != null) {
                newClip = new Area(previousClip);
                newClip.intersect(clipArea);
            }
            graphics.setClip(newClip);
            try {
                super.drawContent(graphics);
            } finally {
                graphics.setClip(previousClip);
            }
        }

        /**
         * Union, sous forme d'{@link Area}, des sous-chemins REMPLIS ({@code
         * PathIf#isFilled()}) du contour geometrique resolu de la forme - voir
         * Javadoc de la classe englobante.
         */
        private Area buildClipArea(Graphics2D graphics) {
            Collection<Outline> outlines = computeResolvedOutlines(graphics);
            Area area = null;
            for (Outline outline : outlines) {
                PathIf path = outline.getPath();
                if (path == null || !path.isFilled()) {
                    continue;
                }
                Area outlineArea = new Area(outline.getOutline());
                if (area == null) {
                    area = outlineArea;
                } else {
                    area.add(outlineArea);
                }
            }
            return area;
        }

        /**
         * Reimplementation de {@code DrawSimpleShape.computeOutlines
         * (Graphics2D)} (POI 5.2.5) utilisant {@link #resolvedGeometry} -
         * deja resolue par heritage de placeholder si necessaire - a la place
         * de {@code getShape().getGeometry()}. L'ancre utilisee reste celle,
         * locale, de l'image elle-meme ({@code getAnchor(graphics, ps)}) :
         * seule la geometrie (le contour) peut venir d'un placeholder
         * herite, jamais la position/taille, toujours definie localement sur
         * le {@code <p:pic>} de la slide en pratique.
         */
        private Collection<Outline> computeResolvedOutlines(Graphics2D graphics) {
            List<Outline> lst = new ArrayList<>();
            if (resolvedGeometry == null) {
                return lst;
            }

            PictureShape<?, ?> ps = getShape();
            Rectangle2D anchor = getAnchor(graphics, ps);
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

                Context ctx = new Context(resolvedGeometry, pathAnchor, ps);
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
}
