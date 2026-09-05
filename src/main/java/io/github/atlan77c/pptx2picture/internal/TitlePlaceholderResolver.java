package io.github.atlan77c.pptx2picture.internal;

import org.apache.poi.sl.usermodel.Placeholder;
import org.apache.poi.xslf.model.PropertyFetcher;
import org.apache.poi.xslf.usermodel.XSLFShape;

/**
 * Resout le type de placeholder OOXML reellement applicable a une forme POI,
 * en suivant si besoin la chaine d'heritage de placeholder PowerPoint (slide
 * -&gt; mise en page -&gt; masque) - notamment pour reconnaitre un titre
 * "centre" ({@code ctrTitle}) dont le type n'est declare qu'au niveau de la
 * mise en page, jamais redeclare localement sur la slide.
 *
 * <p><b>Pourquoi {@code XSLFShape.getPlaceholder()} de POI ne suffit pas
 * seul</b> (verifie dans le code source d'Apache POI 5.2.5, {@code
 * XSLFPlaceholderDetails.getPlaceholder()}) :
 * <pre>{@code
 * public Placeholder getPlaceholder() {
 *     final CTPlaceholder ph = getCTPlaceholder(false);
 *     if (ph == null || !(ph.isSetType() || ph.isSetIdx())) {
 *         return null;
 *     }
 *     return Placeholder.lookupOoxml(ph.getType().intValue());
 * }
 * }</pre>
 * Cette methode n'inspecte QUE le {@code <p:ph>} propre a la forme, sur la
 * SLIDE elle-meme - jamais la mise en page ni le masque. Or en OOXML, une
 * forme placeholder peut parfaitement ne porter localement QUE son {@code
 * idx} (ex. {@code <p:ph idx="0"/>}), le TYPE (title, ctrTitle, body...)
 * n'etant declare qu'une seule fois, au niveau de la mise en page - motif
 * courant, en particulier pour les diapositives utilisant un gabarit "Titre
 * centre". Dans ce cas, {@code ph.isSetType()} est faux mais {@code
 * ph.isSetIdx()} est vrai : le test ci-dessus ne renvoie donc PAS {@code
 * null} comme on pourrait s'y attendre, mais {@code ph.getType().intValue()}
 * - qui, XMLBeans appliquant la valeur par defaut du schema XSD pour un
 * attribut non defini, vaut silencieusement {@code body}. Une forme de titre
 * centre dont le type n'est declare qu'au niveau de la mise en page se voit
 * donc a tort resolue en {@link Placeholder#BODY} par POI - exactement le
 * meme defaut de resolution d'heritage, sur la meme classe de propriete, que
 * celui deja identifie et corrige pour {@code getTextAutofit()} (voir la
 * Javadoc de {@link OverflowAwareTextFitter}, section "Elargissement cible
 * ...") et pour {@code getShapeType()}/{@code getGeometry()} (voir {@link
 * PlaceholderGeometryResolver}).
 *
 * <p><b>Correctif</b> : {@link #resolve} suit la chaine d'heritage via {@code
 * XSLFShape.fetchShapeProperty(PropertyFetcher)} ({@code public}, deja
 * reutilise dans ce paquetage par {@link PlaceholderGeometryResolver} pour la
 * meme raison) : a chaque maillon visite (slide, puis mise en page, puis
 * masque), {@code getPlaceholder()} de POI est appele sur CE maillon - s'il
 * resout {@link Placeholder#TITLE} ou {@link Placeholder#CENTERED_TITLE}, la
 * resolution s'arrete la (le maillon en question declare bien explicitement
 * ce type, aucune ambiguite possible a cet endroit). Sinon (resultat {@code
 * null} OU {@link Placeholder#BODY} - les deux valeurs que ce defaut de POI
 * peut produire indistinctement, voir ci-dessus), on continue vers le maillon
 * suivant plutot que de conclure prematurement. La chaine se termine
 * naturellement au masque, ou le type est toujours declare explicitement.
 */
final class TitlePlaceholderResolver {

    private TitlePlaceholderResolver() {
    }

    /**
     * Resout le type de placeholder reellement applicable a {@code shape} -
     * voir la Javadoc de la classe. Renvoie {@code null} si aucun type ne
     * peut etre resolu (forme non-placeholder, ou chaine d'heritage
     * entierement parcourue sans jamais rencontrer de type explicite).
     */
    static Placeholder resolve(XSLFShape shape) {
        PropertyFetcher<Placeholder> fetcher = new PropertyFetcher<Placeholder>() {
            @Override
            public boolean fetch(XSLFShape candidate) {
                Placeholder placeholder = candidate.getPlaceholder();
                if (placeholder == Placeholder.TITLE || placeholder == Placeholder.CENTERED_TITLE) {
                    setValue(placeholder);
                    return true;
                }
                // null ou BODY (potentiellement un repli par defaut - voir Javadoc de
                // la classe) : rien de fiable a ce maillon, fetchShapeProperty passe au
                // suivant (mise en page, puis masque).
                return false;
            }
        };
        shape.fetchShapeProperty(fetcher);
        return fetcher.isSet() ? fetcher.getValue() : null;
    }

    /**
     * {@code true} si le type resolu (voir {@link #resolve}) designe un
     * titre - standard ({@link Placeholder#TITLE}) ou centre ({@link
     * Placeholder#CENTERED_TITLE}).
     */
    static boolean isTitlePlaceholder(XSLFShape shape) {
        Placeholder resolved = resolve(shape);
        return resolved == Placeholder.TITLE || resolved == Placeholder.CENTERED_TITLE;
    }
}
