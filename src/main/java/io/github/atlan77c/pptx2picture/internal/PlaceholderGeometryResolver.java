package io.github.atlan77c.pptx2picture.internal;

import org.apache.poi.sl.draw.geom.CustomGeometry;
import org.apache.poi.sl.draw.geom.PresetGeometries;
import org.apache.poi.sl.usermodel.ShapeType;
import org.apache.poi.sl.usermodel.SimpleShape;
import org.apache.poi.xslf.model.PropertyFetcher;
import org.apache.poi.xslf.usermodel.XSLFShape;

/**
 * Resout la geometrie (et, quand elle existe, le type de forme preregle)
 * reellement applicable a une forme POI, en suivant si besoin la chaine
 * d'heritage de placeholder PowerPoint (slide -&gt; mise en page -&gt;
 * masque) quand la forme elle-meme ne declare localement ni {@code
 * <a:prstGeom>} ni {@code <a:custGeom>}.
 *
 * <p>Facture, a partir de {@link PictureGeometryClipFixer} (ou ce mecanisme a
 * ete developpe et documente en detail la premiere fois - voir sa Javadoc,
 * sections "Correction du 2026-08-26" et "Correction du 2026-08-29", pour
 * l'historique complet des deux ecarts corriges), afin d'etre reutilise tel
 * quel par {@link AutoShapeGeometryFixer} : le meme calcul de geometrie
 * "reellement applicable" est necessaire pour deux problemes distincts -
 * decouper une image (dont le dessin standard de POI ignore completement le
 * contour geometrique) et calculer le bon contour de remplissage d'une forme
 * ordinaire (texte, forme automatique...) - mais le calcul lui-meme, avec ses
 * deux pieges deja resolus ci-dessous, est identique dans les deux cas.
 *
 * <p><b>Pourquoi {@code getShapeType()}/{@code getGeometry()} de POI ne
 * suffisent pas seuls</b> (verifie dans le code source d'Apache POI 5.2.5,
 * {@code XSLFSimpleShape}) :
 * <ul>
 * <li>Les deux methodes n'inspectent QUE le {@code <p:spPr>} propre a la
 * forme - jamais l'heritage de placeholder (slide -&gt; layout -&gt; master).
 * Une forme de decoupe/remplissage definie uniquement dans la mise en page
 * (cas courant des espaces reserves) est donc invisible pour un appel direct.</li>
 * <li>{@code getShapeType()} ne lit en plus QUE {@code <a:prstGeom>} - elle
 * renvoie {@code null} pour une forme dont la geometrie est un {@code
 * <a:custGeom>} (forme libre dessinee a la main), MEME quand ce {@code
 * custGeom} est defini localement sur le maillon examine.</li>
 * </ul>
 *
 * <p><b>Correctif</b> : {@link #resolve} suit la chaine d'heritage via {@code
 * XSLFShape.fetchShapeProperty(PropertyFetcher)} ({@code public}, deja
 * utilise en interne par POI pour d'autres proprietes - couleur de
 * remplissage, de trait...), et pour chaque maillon distingue trois cas via
 * {@code getGeometry()} plutot que {@code getShapeType()} seul : {@code
 * prstGeom} local (preset nomme), {@code custGeom} local (forme libre, sans
 * preset nomme), ou aucun des deux - ce dernier cas se reconnait en comparant
 * PAR IDENTITE le resultat de {@code getGeometry()} a l'instance de repli
 * partagee que {@code XSLFSimpleShape.getGeometry()} renvoie systematiquement
 * en l'absence de geometrie locale ({@code PresetGeometries.getInstance()}
 * est un singleton charge une seule fois par le JVM, {@code get("rect")}
 * renvoie donc toujours la meme reference - verifie empiriquement avant
 * d'ecrire ce correctif).
 */
final class PlaceholderGeometryResolver {

    /**
     * Instance de repli partagee que {@code XSLFSimpleShape.getGeometry()}
     * renvoie systematiquement quand une forme ne declare localement ni
     * {@code <a:prstGeom>} ni {@code <a:custGeom>} - voir la Javadoc de la
     * classe. Comparee par IDENTITE (pas {@code equals}) dans {@link
     * #resolve} pour distinguer ce repli implicite d'une veritable forme
     * libre locale.
     */
    private static final CustomGeometry IMPLICIT_RECT_GEOMETRY = PresetGeometries.getInstance().get("rect");

    private PlaceholderGeometryResolver() {
    }

    /**
     * Resout la geometrie et le type de forme reellement applicables a
     * {@code shape} - voir la Javadoc de la classe pour le detail du
     * mecanisme et des deux pieges evites.
     */
    static ResolvedGeometry resolve(SimpleShape<?, ?> shape) {
        // Detour par Object avant l'instanceof/cast vers XSLFShape : les parametres
        // generiques "auto-bornes" (S extends Shape<S,P>, ...) de SimpleShape<S,P>,
        // une fois captures depuis un joker <?, ?>, empechent javac d'accepter un
        // instanceof/cast direct vers XSLFShape ("incompatible types : cannot be
        // converted", constate a la compilation reelle). Object n'ayant aucun
        // parametre generique, ce detour contourne cette limitation du compilateur
        // sans changer le comportement a l'execution.
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
                if (type != null) {
                    // <a:prstGeom> local sur ce maillon (RECT explicite compris).
                    setValue(new ResolvedGeometry(type, simpleShape.getGeometry()));
                    return true;
                }
                // getShapeType() ne lit jamais <a:custGeom> : un retour null ici ne
                // veut pas forcement dire "rien sur ce maillon", ca peut aussi etre
                // une forme libre locale. getGeometry(), lui, distingue les deux -
                // sauf qu'il ne renvoie pas non plus explicitement "rien trouve", il
                // retombe silencieusement sur l'instance de repli partagee du preset
                // "rect" dans ce cas. On compare donc par IDENTITE a cette instance
                // pour savoir si on a affaire a une veritable forme libre locale.
                CustomGeometry geometry = simpleShape.getGeometry();
                if (geometry != null && geometry != IMPLICIT_RECT_GEOMETRY) {
                    // Forme libre locale : pas de ShapeType associe (aucun nom de
                    // preset n'existe pour un custGeom), mais bien une geometrie
                    // exploitable - voir ResolvedGeometry.qualifiesForClipping().
                    setValue(new ResolvedGeometry(null, geometry));
                    return true;
                }
                // Vraiment rien a ce maillon - fetchShapeProperty passe au suivant
                // (layout, puis master).
                return false;
            }
        };
        xslfShape.fetchShapeProperty(fetcher);
        return fetcher.isSet() ? fetcher.getValue() : new ResolvedGeometry(null, null);
    }

    /**
     * Resultat de {@link #resolve} : le type de forme et la geometrie
     * resolus, provenant du meme maillon de la chaine d'heritage (jamais
     * l'un de la slide et l'autre de la mise en page).
     */
    static final class ResolvedGeometry {
        final ShapeType type;
        final CustomGeometry geometry;

        ResolvedGeometry(ShapeType type, CustomGeometry geometry) {
            this.type = type;
            this.geometry = geometry;
        }

        /**
         * {@code true} si cette resolution designe une geometrie non
         * rectangulaire qu'il faut prendre en compte (decoupe d'image ou
         * contour de remplissage - voir les utilisateurs de cette classe).
         * {@code type} peut valoir {@code null} ici tout en qualifiant :
         * c'est le cas normal d'une forme libre ({@code custGeom}) qui n'a
         * par nature aucun {@link ShapeType} de preset associe - {@code
         * geometry} est alors le seul signal fiable. Seul un {@code type}
         * explicitement egal a {@link ShapeType#RECT} exclut la prise en
         * compte (un rectangle ne change rien au rendu standard de POI).
         */
        boolean qualifiesForClipping() {
            if (geometry == null || type == ShapeType.RECT) {
                return false;
            }
            return true;
        }
    }
}
