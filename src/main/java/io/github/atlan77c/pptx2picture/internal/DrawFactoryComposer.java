package io.github.atlan77c.pptx2picture.internal;

import org.apache.poi.sl.draw.DrawConnectorShape;
import org.apache.poi.sl.draw.DrawFactory;
import org.apache.poi.sl.draw.DrawPictureShape;
import org.apache.poi.sl.draw.DrawTextShape;
import org.apache.poi.sl.draw.Drawable;
import org.apache.poi.sl.draw.geom.CustomGeometry;
import org.apache.poi.sl.usermodel.ConnectorShape;
import org.apache.poi.sl.usermodel.PictureShape;
import org.apache.poi.sl.usermodel.TextShape;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.awt.Graphics2D;

/**
 * Compose, en un seul {@link DrawFactory}, les correctifs de rendu de ce
 * paquetage bases sur le hint {@code Drawable.DRAW_FACTORY}
 * ({@link ConnectorArrowFixer}, pointes de fleche des connecteurs courbes/
 * coudes ; {@link PictureGeometryClipFixer}, decoupe des images a geometrie
 * non rectangulaire ; {@link PictureAlphaModFixer}, transparence {@code
 * alphaModFix} des images).
 *
 * <p><b>Pourquoi une composition est necessaire</b> : {@code
 * Drawable.DRAW_FACTORY} est un hint de rendu a valeur UNIQUE - {@code
 * graphics.setRenderingHint(Drawable.DRAW_FACTORY, ...)} remplace purement et
 * simplement toute valeur precedente, il ne s'empile pas. Installer les
 * correctifs independamment, l'un apres l'autre (ex. {@code
 * ConnectorArrowFixer.installBeforeDraw(graphics)} puis {@code
 * PictureGeometryClipFixer.installBeforeDraw(graphics)}), ecraserait donc
 * silencieusement les precedents : seul le dernier correctif installe
 * s'appliquerait (chacun ne delegue au comportement standard de POI, via
 * {@code super.getDrawable(...)}, que pour les types de forme qu'il ne gere
 * pas lui-meme - jamais vers le {@code DrawFactory} installe juste avant
 * lui).
 *
 * <p>Cette classe cree un {@code DrawFactory} unique qui delegue chaque type
 * de forme concerne au(x) correctif(s) correspondant(s), et l'installe comme
 * hint unique - voir {@link io.github.atlan77c.pptx2picture.PptxSlideRenderer}.
 *
 * <p><b>Ajout du 2026-08-29</b> : {@link AutoShapeGeometryFixer} (contour de
 * remplissage d'une forme ORDINAIRE - texte, forme automatique - resolu par
 * heritage de placeholder, distinct de {@link PictureGeometryClipFixer} qui
 * ne porte que sur les images - voir sa Javadoc) est desormais compose ici
 * aussi, sur {@code TextShape} plutot que {@code PictureShape}/{@code
 * ConnectorShape} - aucun conflit possible entre les trois, chacun portant
 * sur un type de forme different cote {@code DrawFactory}.
 *
 * <p><b>Ajout du 2026-09-04 - cas particulier de {@code TextShape}</b> :
 * {@link AutoNumberLevelFixer} (numerotation automatique par NIVEAU
 * d'indentation, {@code <a:buAutoNum>} - voir sa Javadoc pour le bug corrige,
 * confirme dans le code source reel de POI) doit s'appliquer a TOUTE forme de
 * texte, y compris celles dont la geometrie resolue par {@code
 * AutoShapeGeometryFixer} reste un simple rectangle (le cas le plus frequent,
 * qu'{@code AutoShapeGeometryFixer} seul laisse au comportement standard de
 * POI - donc SANS correction de numerotation s'il n'etait compose qu'avec
 * lui-meme). Contrairement aux images (deux correctifs INDEPENDANTS qui
 * s'enveloppent l'un l'autre, voir plus bas), un seul objet {@code
 * DrawTextShape} concret est retourne au final : {@code getDrawable(TextShape)}
 * compose donc ici la DECISION d'{@code AutoShapeGeometryFixer} ({@link
 * AutoShapeGeometryFixer#resolveIfQualifies}, la geometrie a utiliser, ou
 * {@code null}) avec le {@code DrawTextShape} produit par {@link
 * AutoNumberLevelFixer#wrap} (qui applique toujours la numerotation, et en
 * prime cette geometrie si elle n'est pas {@code null}) - voir la Javadoc de
 * {@link AutoNumberLevelFixer}, section "Composition avec AutoShapeGeometryFixer".
 *
 * <p><b>Cas particulier des images</b> : {@link PictureGeometryClipFixer} et
 * {@link PictureAlphaModFixer} portent tous deux sur les images ({@code
 * PictureShape}), mais corrigent deux proprietes independantes qui peuvent
 * s'appliquer simultanement a la meme image (une photo peut etre a la fois
 * decoupee selon une forme ET semi-transparente) - ils ne sont donc jamais
 * mis en concurrence l'un de l'autre comme {@code ConnectorArrowFixer} et
 * {@code PictureGeometryClipFixer} le sont entre types de forme differents :
 * {@code PictureAlphaModFixer.wrap(...)} enveloppe systematiquement le
 * resultat de {@code PictureGeometryClipFixer} (ce dernier, ou le
 * comportement standard de POI pour une image rectangulaire) plutot que de
 * s'y substituer - voir la Javadoc de {@link PictureAlphaModFixer}.
 */
public final class DrawFactoryComposer {

    private static final Logger LOG = LoggerFactory.getLogger(DrawFactoryComposer.class);

    private DrawFactoryComposer() {
    }

    /**
     * A appeler juste avant {@code slide.draw(graphics)}, a la place d'un
     * appel individuel a {@code installBeforeDraw} sur chaque correctif -
     * voir Javadoc de la classe.
     *
     * @return la valeur precedente du hint {@code Drawable.DRAW_FACTORY} sur
     * ce {@code Graphics2D} (peut etre {@code null}), a repasser telle quelle
     * a {@link #restoreAfterDraw} une fois {@code slide.draw(graphics)}
     * termine.
     */
    public static Object installBeforeDraw(Graphics2D graphics) {
        Object previous = graphics.getRenderingHint(Drawable.DRAW_FACTORY);
        DrawFactory connectorFactory = ConnectorArrowFixer.createDrawFactory();
        DrawFactory pictureFactory = PictureGeometryClipFixer.createDrawFactory();
        graphics.setRenderingHint(Drawable.DRAW_FACTORY, new DrawFactory() {
            @Override
            public DrawConnectorShape getDrawable(ConnectorShape<?, ?> shape) {
                return connectorFactory.getDrawable(shape);
            }

            @Override
            public DrawPictureShape getDrawable(PictureShape<?, ?> shape) {
                // PictureGeometryClipFixer decide d'abord de la decoupe (ou renvoie le
                // DrawPictureShape standard de POI si l'image est rectangulaire) ;
                // PictureAlphaModFixer enveloppe ENSUITE ce resultat pour appliquer, en
                // plus et independamment, la transparence alphaModFix du blip - voir
                // Javadoc de la classe, section "Cas particulier des images".
                DrawPictureShape geometryAware = pictureFactory.getDrawable(shape);
                return PictureAlphaModFixer.wrap(shape, geometryAware);
            }

            @Override
            public DrawTextShape getDrawable(TextShape<?, ?> shape) {
                // AutoShapeGeometryFixer decide seulement QUELLE geometrie utiliser (null =
                // rectangle standard) ; AutoNumberLevelFixer fournit l'UNIQUE DrawTextShape
                // concret retourne, qui applique a la fois cette geometrie (le cas echeant) et
                // la correction de numerotation par niveau - voir Javadoc de la classe, section
                // "Ajout du 2026-09-04", et Javadoc de AutoNumberLevelFixer.
                CustomGeometry resolvedGeometry = AutoShapeGeometryFixer.resolveIfQualifies(shape);
                return AutoNumberLevelFixer.wrap(shape, resolvedGeometry);
            }
        });
        if (LOG.isDebugEnabled()) {
            LOG.debug("DrawFactory compose (connecteurs + images decoupees selon leur forme + transparence "
                    + "alphaModFix + contour heritee des formes ordinaires + numerotation automatique par "
                    + "niveau) installe");
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
}
