package io.github.atlan77c.pptx2image;

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

    private RenderOptions(Builder builder) {
        this.scale = builder.scale;
        this.fixTextOverflow = builder.fixTextOverflow;
        this.background = builder.background;
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

        public RenderOptions build() {
            return new RenderOptions(this);
        }
    }
}
