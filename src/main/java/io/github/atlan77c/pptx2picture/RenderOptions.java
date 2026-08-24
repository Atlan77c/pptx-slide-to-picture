package io.github.atlan77c.pptx2picture;

import java.awt.Color;
import java.util.Objects;

/**
 * Options de rendu pour {@link PptxSlideRenderer}. Instance immuable ;
 * construite via {@link #builder()}. Les valeurs par defaut (voir chaque
 * methode du {@link Builder}) conviennent a la grande majorite des cas.
 *
 * <pre>{@code
 * RenderOptions options = RenderOptions.builder()
 *     .scale(2.0f)
 *     .build();
 * }</pre>
 */
public final class RenderOptions {

    private final float scale;
    private final boolean fixTextOverflow;
    private final Color background;
    private final OutputFormat format;
    private final float jpegQuality;

    private RenderOptions(Builder builder) {
        this.scale = builder.scale;
        this.fixTextOverflow = builder.fixTextOverflow;
        this.background = builder.background;
        this.format = builder.format;
        this.jpegQuality = builder.jpegQuality;
    }

    /** Facteur d'echelle applique a la taille native de la slide (voir {@link Builder#scale(float)}). */
    public float getScale() {
        return scale;
    }

    /** Indique si le correctif de debordement de texte est actif (voir {@link Builder#fixTextOverflow(boolean)}). */
    public boolean isFixTextOverflow() {
        return fixTextOverflow;
    }

    /** Couleur de fond de l'image produite (voir {@link Builder#background(Color)}). */
    public Color getBackground() {
        return background;
    }

    /**
     * Format de sortie utilise par {@code PptxSlideRenderer.renderSlideToFile(...)}
     * (voir {@link Builder#format(OutputFormat)}). Sans effet sur {@code renderSlide}
     * (toujours un {@link java.awt.image.BufferedImage} raster) ni sur
     * {@code renderSlideAsSvg} (toujours du SVG) - seule la variante "vers fichier"
     * en tient compte pour choisir l'encodage a ecrire.
     */
    public OutputFormat getFormat() {
        return format;
    }

    /**
     * Qualite de compression JPEG, utilisee uniquement lorsque {@link #getFormat()}
     * vaut {@link OutputFormat#JPEG} (voir {@link Builder#jpegQuality(float)}).
     */
    public float getJpegQuality() {
        return jpegQuality;
    }

    public static Builder builder() {
        return new Builder();
    }

    /** Options par defaut : equivalent a {@code RenderOptions.builder().build()}. */
    public static RenderOptions defaults() {
        return builder().build();
    }

    public static final class Builder {
        private float scale = 2.0f;
        private boolean fixTextOverflow = true;
        private Color background = Color.WHITE;
        private OutputFormat format = OutputFormat.PNG;
        private float jpegQuality = 0.92f;

        private Builder() {
        }

        /**
         * Facteur d'echelle applique a la taille native de la slide (en points,
         * generalement 960x540 ou 1280x720). Par defaut {@code 2.0f} : produit une
         * image a une resolution environ 2x superieure a la taille native, pour un
         * rendu net a l'ecran et a l'impression. Doit etre strictement positif.
         */
        public Builder scale(float scale) {
            if (scale <= 0) {
                throw new IllegalArgumentException("scale doit etre strictement positif, recu: " + scale);
            }
            this.scale = scale;
            return this;
        }

        /**
         * Active (par defaut) ou desactive le correctif de debordement de texte
         * (voir la Javadoc de {@link PptxSlideRenderer}, section "Fidelite du rendu
         * et limites connues"). Desactiver ce correctif restitue le comportement
         * brut d'Apache POI : certaines formes de texte peuvent alors deborder
         * visuellement de leur emplacement d'origine.
         */
        public Builder fixTextOverflow(boolean fixTextOverflow) {
            this.fixTextOverflow = fixTextOverflow;
            return this;
        }

        /**
         * Couleur de fond de l'image produite, utilisee derriere le contenu de la
         * slide (les slides pptx sans arriere-plan explicite sont transparentes
         * dans le modele OOXML). Par defaut blanc.
         */
        public Builder background(Color background) {
            this.background = Objects.requireNonNull(background, "background");
            return this;
        }

        /**
         * Format de sortie utilise par {@code PptxSlideRenderer.renderSlideToFile(...)}.
         * Par defaut {@link OutputFormat#PNG}. Voir {@link OutputFormat} pour le
         * detail des formats disponibles et leurs limites respectives.
         */
        public Builder format(OutputFormat format) {
            this.format = Objects.requireNonNull(format, "format");
            return this;
        }

        /**
         * Qualite de compression JPEG, entre {@code 0.0} (taille minimale, qualite
         * la plus degradee) et {@code 1.0} (qualite maximale, quasi sans perte).
         * Par defaut {@code 0.92f}. Sans effet si {@link #format(OutputFormat)}
         * n'est pas {@link OutputFormat#JPEG}.
         */
        public Builder jpegQuality(float jpegQuality) {
            if (jpegQuality < 0f || jpegQuality > 1f) {
                throw new IllegalArgumentException("jpegQuality doit etre compris entre 0.0 et 1.0, recu: " + jpegQuality);
            }
            this.jpegQuality = jpegQuality;
            return this;
        }

        public RenderOptions build() {
            return new RenderOptions(this);
        }
    }
}
