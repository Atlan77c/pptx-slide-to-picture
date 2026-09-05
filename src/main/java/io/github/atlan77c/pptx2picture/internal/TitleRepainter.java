package io.github.atlan77c.pptx2picture.internal;

import org.apache.poi.sl.draw.DrawFactory;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.awt.Graphics2D;

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
 * <p><b>Detection d'un titre</b> : deleguee a {@link TitleDetector} (type de
 * placeholder OOXML herite du masque/layout, avec repli sur le prefixe du nom
 * de la forme pour un titre dissocie de son placeholder - voir sa Javadoc
 * pour le detail complet des deux mecanismes).
 *
 * <p><b>Correction du 2026-08-31 : titre "centre" ({@code ctrTitle}) non
 * reconnu quand son type n'est declare qu'au niveau de la mise en page</b>.
 * Jusque-la, cette classe comparait directement {@code
 * XSLFTextShape.getPlaceholder()} a {@code Placeholder.TITLE}/{@code
 * CENTERED_TITLE} - or cette methode POI ne suit jamais l'heritage de
 * placeholder (slide -&gt; mise en page -&gt; masque, voir la Javadoc de
 * {@link TitlePlaceholderResolver} pour la preuve complete dans le code
 * source de POI) : une forme dont le {@code <p:ph>} local ne porte qu'un
 * {@code idx} (type "centre" declare uniquement dans la mise en page - motif
 * courant du gabarit "Titre centre") se voyait a tort resolue en {@code
 * BODY}, jamais repeinte par cette classe. {@link TitleDetector} s'appuie
 * desormais sur {@link TitlePlaceholderResolver}, qui suit correctement cette
 * chaine d'heritage - correctif sans risque de regression sur les cas deja
 * corrects (type declare localement, resolu des le premier maillon, comme
 * avant).
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
            if (TitleDetector.isTitleShape(shape)) {
                DrawFactory.getInstance(graphics).getDrawable(shape).draw(graphics);
                repainted++;
                if (LOG.isDebugEnabled()) {
                    LOG.debug("{} : titre repeint au premier plan (voir Javadoc de la classe)", shape.getShapeName());
                }
            }
        }
        return repainted;
    }
}
