package io.github.atlan77c.pptx2picture.internal;

import org.apache.poi.sl.draw.DrawConnectorShape;
import org.apache.poi.sl.draw.DrawFactory;
import org.apache.poi.sl.draw.DrawPictureShape;
import org.apache.poi.sl.draw.Drawable;
import org.apache.poi.sl.usermodel.ConnectorShape;
import org.apache.poi.sl.usermodel.PictureShape;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.awt.Graphics2D;

/**
 * Compose, en un seul {@link DrawFactory}, les correctifs de rendu de ce
 * paquetage bases sur le hint {@code Drawable.DRAW_FACTORY}
 * ({@link ConnectorArrowFixer}, pointes de fleche des connecteurs courbes/
 * coudes, et {@link PictureGeometryClipFixer}, decoupe des images a
 * geometrie non rectangulaire).
 *
 * <p><b>Pourquoi une composition est necessaire</b> : {@code
 * Drawable.DRAW_FACTORY} est un hint de rendu a valeur UNIQUE - {@code
 * graphics.setRenderingHint(Drawable.DRAW_FACTORY, ...)} remplace purement et
 * simplement toute valeur precedente, il ne s'empile pas. Installer les deux
 * correctifs independamment, l'un apres l'autre (ex. {@code
 * ConnectorArrowFixer.installBeforeDraw(graphics)} puis {@code
 * PictureGeometryClipFixer.installBeforeDraw(graphics)}), ecraserait donc
 * silencieusement le premier {@code DrawFactory} par le second : le
 * correctif installe en premier cesserait purement et simplement de
 * s'appliquer (chacun ne delegue au comportement standard de POI, via {@code
 * super.getDrawable(...)}, que pour les types de forme qu'il ne gere pas
 * lui-meme - jamais vers le {@code DrawFactory} installe juste avant lui).
 *
 * <p>Cette classe cree un {@code DrawFactory} unique qui delegue chaque type
 * de forme concerne au correctif correspondant, et l'installe comme hint
 * unique - voir {@link io.github.atlan77c.pptx2picture.PptxSlideRenderer}.
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
                return pictureFactory.getDrawable(shape);
            }
        });
        if (LOG.isDebugEnabled()) {
            LOG.debug("DrawFactory compose (connecteurs + images decoupees selon leur forme) installe");
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
