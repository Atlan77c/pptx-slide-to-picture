package io.github.atlan77c.pptx2picture.internal;

import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextShape;

import java.util.List;

/**
 * Detection d'une forme de titre de diapositive, factorisee ici pour etre
 * partagee entre {@link TitleRepainter} (repeindre chaque titre par-dessus le
 * reste de la slide - voir historique complet de ce besoin dans sa Javadoc,
 * ou cette detection a ete developpee la premiere fois) et {@link
 * OverflowAwareTextFitter} (n'elargir un correctif qu'aux diapositives de
 * sommaire/table des matieres, identifiees par le TEXTE de leur titre - voir
 * sa Javadoc, section "Elargissement cible...").
 *
 * <p><b>Deux niveaux de detection</b>, dans cet ordre :
 * <ol>
 * <li>le type de placeholder OOXML de la forme, resolu via {@link
 * TitlePlaceholderResolver} (suit l'heritage slide -&gt; mise en page -&gt;
 * masque, ce que l'API POI ne fait pas nativement - voir sa Javadoc pour le
 * detail, notamment necessaire pour un titre "centre" dont le type n'est
 * declare qu'au niveau de la mise en page) : {@link
 * org.apache.poi.sl.usermodel.Placeholder#TITLE} ou {@link
 * org.apache.poi.sl.usermodel.Placeholder#CENTERED_TITLE} ;</li>
 * <li>a defaut, le prefixe du NOM de la forme (voir {@link
 * #TITLE_NAME_PREFIXES}) - necessaire quand un auteur PowerPoint a dissocie
 * un titre de son placeholder (copier-coller, reorganisation manuelle...) tout
 * en conservant le nom par defaut "Titre N"/"Title N" affiche dans le volet
 * Selection ; cas deja rencontre sur plusieurs diapositives de fichiers reels
 * de ce projet.</li>
 * </ol>
 *
 * <p><b>Limite assumee</b> : seules les formes de premier niveau d'une
 * diapositive sont examinees (pas celles a l'interieur d'un groupe) - un
 * titre n'a, dans les fichiers observes, jamais ete rencontre imbrique dans
 * un groupe.
 */
final class TitleDetector {

    /**
     * Prefixes de nom de forme reconnus comme "titre" quand le type de
     * placeholder n'est pas resolu (voir Javadoc de la classe). Comparaison
     * insensible a la casse, sur le debut du nom uniquement (ex. "Titre 3"
     * correspond au prefixe "Titre", mais pas un futur "Titrefoo").
     *
     * <p><b>Pour ajouter une langue</b> : ajouter le mot "Titre" tel que
     * PowerPoint le genere par defaut, dans cette langue, pour une forme de
     * titre dissociee de son placeholder - verifiable via le volet Selection
     * de PowerPoint ou l'attribut {@code name} de {@code <p:nvSpPr><p:cNvPr>}
     * dans le XML de la forme.
     */
    static final List<String> TITLE_NAME_PREFIXES = List.of(
            "Titre",   // francais
            "Title",   // anglais
            "Título",  // espagnol / portugais (bresilien)
            "Titolo",  // italien
            "Titel"    // allemand
    );

    private TitleDetector() {
    }

    /** {@code true} si {@code shape} est detectee comme un titre - voir Javadoc de la classe. */
    static boolean isTitleShape(XSLFShape shape) {
        if (!(shape instanceof XSLFTextShape)) {
            return false;
        }
        if (TitlePlaceholderResolver.isTitlePlaceholder(shape)) {
            return true;
        }
        return hasTitleNamePrefix(shape.getShapeName());
    }

    private static boolean hasTitleNamePrefix(String shapeName) {
        if (shapeName == null) {
            return false;
        }
        String trimmed = shapeName.trim();
        for (String prefix : TITLE_NAME_PREFIXES) {
            if (trimmed.regionMatches(true, 0, prefix, 0, prefix.length())) {
                int nextIndex = prefix.length();
                // limite de mot : accepte "Titre 3" mais pas un nom qui commence
                // simplement par les memes lettres, ex. un futur "Titrefoo".
                if (trimmed.length() == nextIndex || !Character.isLetter(trimmed.charAt(nextIndex))) {
                    return true;
                }
            }
        }
        return false;
    }

    /**
     * Renvoie la premiere forme de titre trouvee au premier niveau de {@code
     * slide} (voir {@link #isTitleShape}), ou {@code null} si aucune n'est
     * detectee.
     */
    static XSLFTextShape findTitleShape(XSLFSlide slide) {
        for (XSLFShape shape : slide.getShapes()) {
            if (shape instanceof XSLFTextShape && isTitleShape(shape)) {
                return (XSLFTextShape) shape;
            }
        }
        return null;
    }
}
