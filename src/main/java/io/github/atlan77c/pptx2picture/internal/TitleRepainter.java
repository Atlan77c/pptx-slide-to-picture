package io.github.atlan77c.pptx2picture.internal;

import org.apache.poi.sl.draw.DrawFactory;
import org.apache.poi.sl.usermodel.Placeholder;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.awt.Graphics2D;
import java.util.List;

/**
 * Corrige un ecart de fidelite decouvert sur un fichier reel : un titre
 * partiellement recouvert par une image ou une forme voisine peinte apres lui
 * dans l'ordre du document, alors que dans PowerPoint le titre reste
 * entierement lisible.
 *
 * <p><b>Analyse</b> : l'ordre de peinture (ordre des formes dans le XML,
 * {@code spTree}) est identique entre PowerPoint et Apache POI - ce n'est
 * donc pas un probleme d'ordre de premier plan/arriere-plan mal reproduit.
 * Le chevauchement lui-meme preexiste dans le fichier source (l'ancrage du
 * titre et celui de la forme voisine se recouvrent deja legerement, avant
 * meme tout rendu). Ce qui "s'aggrave" au rendu est, comme pour
 * {@link OverflowAwareTextFitter}, l'ecart recurrent de metrique de police
 * entre Java2D/AWT et le moteur de rendu de PowerPoint (jusqu'a ~30-35%
 * observe ailleurs dans ce projet) : meme quand aucun debordement n'est
 * mesure (la boite du titre est largement assez haute pour une seule ligne
 * de texte, donc {@link OverflowAwareTextFitter} reste totalement passif sur
 * cette forme), l'etendue verticale reellement occupee par le texte rendu,
 * sous un ancrage Haut, peut etre plus grande chez nous que chez PowerPoint -
 * poussant davantage de glyphes dans la zone deja recouverte par la forme
 * peinte apres le titre dans l'ordre du document.
 *
 * <p><b>Correctif retenu</b> : repeindre systematiquement chaque forme de
 * titre PAR-DESSUS le reste de la slide, une fois {@code slide.draw(graphics)}
 * termine - sans toucher a l'ordre des formes dans le XML (evite tout risque
 * de desynchroniser l'etat interne de POI, comme documente dans l'historique
 * de {@link SymbolFontRunFixer}). Applique sans aucune condition de
 * chevauchement detecte, y compris sur les slides sans probleme : repeindre
 * un titre deja correctement affiche produit exactement les memes pixels
 * (aucun effet visible), donc aucun risque de regression sur les slides bien
 * formees.
 *
 * <p><b>Detection d'un titre</b> : la maniere fiable, quand elle est
 * disponible, est le type de placeholder OOXML herite du masque/layout de la
 * diapositive ({@code <p:ph type="title"/>} ou {@code type="ctrTitle"}). Mais
 * un auteur peut, dans PowerPoint, dissocier un titre de son placeholder (le
 * remplacer par une simple zone de texte, ex. via un copier-coller ou une
 * reorganisation manuelle) tout en conservant le nom "Titre N" affiche dans
 * le volet Selection - ce cas a ete observe sur plusieurs slides d'un fichier
 * reel (un titre dissocie de son placeholder cote a cote avec un autre resté
 * lie a son placeholder). Dans ce cas, on se rabat sur le prefixe du nom de la
 * forme, teste insensible a la casse contre une liste de prefixes connus dans
 * plusieurs langues (voir {@link #TITLE_NAME_PREFIXES}).
 *
 * <p><b>Limite assumee</b> : seules les formes de premier niveau de la slide
 * sont examinees (pas celles a l'interieur d'un groupe). Repeindre une forme
 * situee dans un groupe necessiterait de rejouer la transformation
 * geometrique de ce groupe pour la positionner correctement, ce qui n'a pas
 * ete rencontre en pratique : un titre est, dans les fichiers observes,
 * toujours soit un veritable placeholder, soit une forme de premier niveau
 * nommee "Titre N" (jamais imbriquee dans un groupe).
 */
public final class TitleRepainter {

    private static final Logger LOG = LoggerFactory.getLogger(TitleRepainter.class);

    /**
     * Prefixes de nom de forme reconnus comme "titre" quand le lien de
     * placeholder OOXML est absent (voir Javadoc de la classe, section
     * "Detection d'un titre"). Comparaison insensible a la casse, sur le
     * debut du nom uniquement (ex. "Titre 3" correspond au prefixe "Titre").
     *
     * <p><b>Pour ajouter une langue</b> : ajouter dans cette liste le mot
     * "Titre" tel que PowerPoint le genere par defaut, dans cette langue,
     * pour une forme de titre dissociee de son placeholder. On peut le
     * verifier soit via le volet Selection de PowerPoint (nom affiche pour
     * une telle forme dans un fichier cree avec la langue voulue), soit en
     * inspectant directement l'attribut {@code name} de l'element
     * {@code <p:nvSpPr><p:cNvPr>} dans le XML de la forme (fichier
     * {@code .pptx} renomme en {@code .zip}, dossier {@code ppt/slides/}).
     */
    private static final List<String> TITLE_NAME_PREFIXES = List.of(
            "Titre",   // francais
            "Title",   // anglais
            "Título",  // espagnol / portugais (bresilien)
            "Titolo",  // italien
            "Titel"    // allemand
    );

    private TitleRepainter() {
    }

    /**
     * Repeint, par-dessus le rendu deja effectue par {@code slide.draw(graphics)},
     * chaque forme de titre detectee (voir Javadoc de la classe). A appeler
     * APRES {@code slide.draw(graphics)} - c'est l'inverse des autres
     * correctifs de ce paquetage, qui modifient l'etat des formes AVANT le
     * dessin de la slide.
     *
     * @return le nombre de titres effectivement repeints.
     */
    public static int repaintTitles(XSLFSlide slide, Graphics2D graphics) {
        int repainted = 0;
        for (XSLFShape shape : slide.getShapes()) {
            if (isTitleShape(shape)) {
                DrawFactory.getInstance(graphics).getDrawable(shape).draw(graphics);
                repainted++;
                if (LOG.isDebugEnabled()) {
                    LOG.debug("{} : titre repeint au premier plan (voir Javadoc de la classe)", shape.getShapeName());
                }
            }
        }
        return repainted;
    }

    private static boolean isTitleShape(XSLFShape shape) {
        if (!(shape instanceof XSLFTextShape)) {
            return false;
        }
        XSLFTextShape ts = (XSLFTextShape) shape;
        Placeholder placeholder = ts.getPlaceholder();
        if (placeholder == Placeholder.TITLE || placeholder == Placeholder.CENTERED_TITLE) {
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
}
