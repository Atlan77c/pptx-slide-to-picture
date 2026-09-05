package io.github.atlan77c.pptx2picture.internal;

import org.apache.poi.sl.usermodel.TextShape;
import org.apache.poi.sl.usermodel.VerticalAlignment;
import org.apache.poi.xslf.usermodel.XSLFGroupShape;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextBox;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.awt.Graphics2D;
import java.awt.geom.Rectangle2D;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Corrige l'angle mort laisse par {@link OverflowAwareTextFitter} : ce
 * dernier ne retrecit une forme {@code noAutofit} que si son texte deborde de
 * <em>sa propre</em> ancre ET que ce debordement chevauche une autre forme.
 * Il ne fait rien quand le texte mesure tient deja dans sa propre ancre - meme
 * si, a l'interieur de cette ancre (volontairement surdimensionnee pour
 * accueillir une forme voisine posee par-dessus, comme un post-it ou une
 * annotation), ce texte s'etend reellement assez bas pour toucher cette forme
 * voisine independante.
 *
 * <p><b>Cas reel a l'origine de ce correctif</b> (slide 10 du meme fichier que
 * {@link OversizedWhitespaceRunFixer}) : une boite "Text Placeholder 6"
 * (id 14, ancre haute de ~363pt) empile six paragraphes de statistiques
 * ("~3 jours...", "~40 %...", "44 %..."). Le texte mesure (quelques 250pt)
 * tient largement dans les 363pt de l'ancre - {@link OverflowAwareTextFitter}
 * ne se declenche donc jamais. Une annotation independante ("ZoneTexte 3",
 * "Nos échanges sont simplifiés") est positionnee par l'auteur a l'interieur
 * de cette meme ancre, dans l'espace que PowerPoint laisse vide sous le
 * dernier paragraphe. Confirme par l'utilisateur par comparaison directe avec
 * une capture PowerPoint : chez POI, a cause de la surestimation Java2D/AWT
 * deja documentee ({@link OverflowAwareTextFitter}, {@link
 * OversizedWhitespaceRunFixer}), le bloc de texte s'etend en realite plus bas
 * que chez PowerPoint et vient chevaucher cette annotation - sans jamais
 * depasser sa propre ancre.
 *
 * <h2>2e variante (2026-08-28) : espaceurs vides devant une annotation, slide 17</h2>
 * <p>Meme fichier, slide 17 : la forme "Espace réservé du texte 2" empile 11
 * paragraphes a 16pt - trois paragraphes de texte visible, entrecoupes de SIX
 * paragraphes ENTIEREMENT VIDES (espaceurs, {@code endParaRPr} seul, aucun
 * {@code <a:r>}) volontairement places pour faire atterrir chaque paragraphe
 * visible dans un espace libre precis, pendant que deux annotations opaques
 * ("ZoneTexte 3", "ZoneTexte 9", dessinees par-dessus) couvrent le reste. Le
 * paragraphe "Il a été éclairé par les facteurs de décision suivants" (le
 * 10e) tient largement au-dessus de "ZoneTexte 9" chez PowerPoint, mais la
 * meme surestimation Java2D/AWT, cumulee sur les six paragraphes vides qui le
 * precedent, le fait glisser assez bas pour chevaucher "ZoneTexte 9" chez
 * nous.
 *
 * <p><b>Pourquoi la version initiale de ce correctif ne l'attrapait pas</b> :
 * {@code computeOccupiedZone()} calcule UNE SEULE zone rectangulaire pour
 * toute la forme, du haut de l'ancre jusqu'a la hauteur totale mesuree
 * (espaceurs vides compris). Pour cette forme precise, cette zone globale
 * chevauche par construction "ZoneTexte 3" ET "ZoneTexte 9" - et ce, QUE LE
 * RENDU SOIT CORRECT OU BUGUE, puisque c'est justement le principe du motif
 * "placeholder geant + espaceurs vides + annotations posees dans les vides" :
 * le bloc de texte est SUPPOSE chevaucher geometriquement les ancres des
 * annotations, seul du blanc devant s'y trouver reellement. Faire disparaitre
 * ce chevauchement "forme entiere" aurait exige de compresser tout le bloc a
 * moins de 150pt de hauteur totale (11 paragraphes) - ce qui aurait rendu tout
 * le texte visible illisible. Le correctif epuisait donc ses deux leviers
 * sans succes et abandonnait : la forme n'etait jamais modifiee, et le vrai
 * defaut (un seul paragraphe, precisement) restait entier.
 *
 * <p><b>Correction retenue</b> : le MECANISME de correction (les deux leviers
 * ci-dessous, appliques uniformement a TOUTE la forme) et son PERIMETRE
 * (formes dont le texte tient dans sa propre ancre) restent inchanges - seul
 * le TEST DE COLLISION change de granularite, de "forme entiere" a
 * "paragraphe par paragraphe" : on ne considere desormais qu'il y a un
 * probleme que si au moins un paragraphe VISIBLE (texte non entierement
 * blanc) chevauche individuellement une autre forme - les paragraphes
 * entierement vides (espaceurs) sont ignores dans ce test, meme s'ils
 * occupent geometriquement un espace qui chevauche une annotation, puisqu'ils
 * ne peignent aucun glyphe. Cette nouvelle condition est plus FAIBLE (donc
 * plus facile a satisfaire) que l'ancienne "toute la forme, espaceurs
 * compris" : si l'ancien test reussissait a resoudre un cas, le nouveau y
 * reussit forcement aussi (la hauteur totale mesuree majore toujours la
 * position de n'importe quel paragraphe individuel) - la generalisation ne
 * peut donc jamais faire regresser le cas d'origine (slide 10, section
 * "2e variante" ci-dessus), voir {@code
 * NeighborShapeOverlapFixerTest#fixNeighborOverlaps_resolvesCollision_whenNeighbourWithinOwnGenerousAnchor}.
 * Pour le nouveau cas (slide 17), elle permet en revanche de s'arreter des
 * que le SEUL paragraphe reellement concerne cesse de chevaucher, sans avoir
 * a satisfaire l'exigence, hors de portee, de degager la totalite du bloc.
 *
 * <p><b>Obstacle technique et technique retenue pour mesurer un paragraphe
 * individuellement</b> : verifie dans le source POI 5.2.5 ({@code
 * XSLFTextShape}, {@code DrawTextShape}, {@code DrawTextParagraph}, miroir
 * GitHub officiel) - il n'existe AUCUNE API publique pour mesurer la
 * hauteur/position d'un paragraphe pris isolement. {@code
 * XSLFTextShape#getTextHeight(Graphics2D)} delegue a un dry-run global
 * ({@code DrawTextShape#drawParagraphs}, {@code package-private}) qui calcule
 * bien la position de chaque paragraphe en interne ({@code
 * DrawTextParagraph#getY()}, public), mais {@code breakText()} - qui peuple
 * les lignes et calcule la hauteur reelle - est {@code protected}, donc
 * inutilisable depuis ce code. {@link ParagraphMeasurer} contourne cela en
 * construisant, une seule fois par forme analysee, une COPIE DE MESURE
 * jetable (un {@link XSLFTextBox} ajoute au meme slide puis retire juste
 * apres, jamais dessine) et en tronquant/mesurant/reconstruisant
 * progressivement CETTE copie (jamais la forme source) via l'API publique
 * deja utilisee partout ailleurs dans ce projet ({@code addNewTextParagraph},
 * {@code removeTextParagraph}, {@code getTextHeight()}) - la hauteur cumulee
 * des k premiers paragraphes, mesuree ainsi, donne par difference la
 * contribution individuelle de chaque paragraphe. Ne mesurer QUE la copie
 * jetable (jamais la forme source) evite tout risque de perdre l'identite
 * des objets {@link XSLFTextParagraph}/{@link XSLFTextRun} de la forme
 * source, dont ce correctif (comme {@link OverflowAwareTextFitter}) depend
 * pour restaurer fidelement les valeurs d'origine en cas d'echec. La copie
 * reproduit uniquement ce qui influence la hauteur mesuree par Java2D :
 * largeur et marges internes de la forme, interligne de chaque paragraphe,
 * texte/taille/police de chaque run - la police est copiee via {@code
 * XSLFTextRun#getFontFamily()} (DEJA RESOLUE, ex. "Calibri" et non
 * "+mn-lt") plutot que la reference de theme brute, pour que la copie de
 * mesure reste fiable independamment de son propre theme. Un paragraphe
 * ENTIEREMENT VIDE (aucun {@code <a:r>}, comme les espaceurs du slide 17) est
 * reproduit par un unique run fantome contenant un espace, a la taille {@code
 * XSLFTextParagraph#getDefaultFontSize()} du paragraphe source - exactement
 * le mecanisme que POI utilise lui-meme en interne ({@code
 * DrawTextParagraph#getAttributedString()}) pour ne jamais mesurer une
 * hauteur nulle sur un paragraphe vide.
 *
 * <p><b>Strategie a deux leviers, dans cet ordre</b> (inchangee, appliquee a
 * TOUTE la forme comme avant) :
 * <ol>
 *   <li><b>Reduire l'interligne</b> ({@code <a:lnSpc><a:spcPct>}) de tous les
 *   paragraphes eligibles de la forme, de facon bornee (voir {@link
 *   #SPACING_MIN_FACTOR}). Contrairement au retrecissement de police, ceci
 *   laisse le texte a sa taille visible d'origine. C'est aussi le levier le
 *   plus proche de la piste "compensation directe de {@code
 *   DrawTextFragment.getHeight()}" deja tentee et abandonnee par ce projet
 *   (voir Javadoc de {@link OverflowAwareTextFitter}) : reduire l'interligne
 *   reduit le pas d'avancement entre deux lignes SANS reduire l'encombrement
 *   reel (ascendant+descendant) de chaque ligne, ce qui peut en theorie faire
 *   chevaucher deux lignes d'un meme paragraphe si on reduit trop. C'est pour
 *   cette raison que la reduction est bornee de facon conservatrice plutot
 *   que poussee jusqu'a disparition complete de la collision.</li>
 *   <li><b>Repli sur le retrecissement de police</b> (meme mecanique que
 *   {@link OverflowAwareTextFitter}, sure par construction puisqu'elle reduit
 *   glyphes et interlignage ensemble et proportionnellement) si la reduction
 *   d'interligne, une fois a sa limite, n'a pas suffi a faire disparaitre la
 *   collision.</li>
 * </ol>
 *
 * <p><b>Justification de la limite de reduction d'interligne</b> ({@link
 * #SPACING_MIN_FACTOR} = 0,85) : {@code DrawTextFragment.getLeading()}
 * (code source POI 5.2.5) retombe sur {@code (ascendant+descendant) * 0.15}
 * des que {@code TextLayout.getLeading()} renvoie 0 - ce qui est le cas
 * courant. Cette part de "interligne intrinseque" (~13% de la hauteur totale
 * ascendant+descendant+interligne dans ce cas courant) est un espace de
 * respiration deliberement ajoute au-dela de l'encombrement reel des glyphes
 * : la reduire ne devrait donc, dans ce cas courant, jamais faire toucher les
 * glyphes de deux lignes consecutives entre eux. Se limiter a 85% (donc ne
 * jamais reduire de plus de 15 points de pourcentage) reste a l'interieur de
 * cette marge avec une petite reserve de securite. <b>Limite assumee</b> :
 * ceci reste un raisonnement analytique, pas une garantie verifiee au cas par
 * cas - {@code TextLayout.getLeading()} peut renvoyer une valeur non nulle
 * selon la police (auquel cas la part reellement "gratuite" differe de 13%),
 * et l'API publique de POI n'expose pas de decomposition ligne par ligne
 * permettant de le verifier programmatiquement ici. D'ou une limite
 * volontairement conservatrice plutot qu'une reduction poussee jusqu'a
 * disparition totale de la collision comme pour la police.
 *
 * <p><b>Complementaire, jamais redondant, avec {@link
 * OverflowAwareTextFitter}</b> : ce correctif ignore toute forme dont le
 * texte mesure depasse deja sa propre ancre (`textHeight > anchor.getHeight()`)
 * - ce cas reste entierement gere par {@link OverflowAwareTextFitter}, qui
 * doit s'executer avant (voir ordre d'appel dans {@code PptxSlideRenderer}).
 *
 * <h2>Formes exemptees de l'elargissement sommaire (2026-08-31)</h2>
 * <p>Exception a la regle ci-dessus, decouverte sur le meme fichier "Refonte
 * BEL" (slide 2, sommaire) que {@link OverflowAwareTextFitter} : depuis le
 * 2026-08-31, celui-ci ne retrecit plus JAMAIS une forme relevant de son
 * elargissement sommaire (voir sa Javadoc, section "Elargissement cible...")
 * - y compris quand le texte de cette forme, a taille pleine, depasse sa
 * PROPRE ancre (motif frequent pour un item de sommaire : ancre dimensionnee
 * pour une seule ligne, texte reel passant sur deux). Un tel cas, s'il
 * chevauche en plus une forme voisine, n'est alors plus corrige par
 * personne : ni par {@link OverflowAwareTextFitter} (qui l'exempte
 * deliberement), ni par ce correctif (dont la regle ci-dessus l'ecarte au
 * profit d'{@link OverflowAwareTextFitter}, precisement parce qu'il etait
 * cense s'en charger). Cas reel : l'item "Contexte – Rappel de la vision
 * cible" du sommaire, sur deux lignes a taille pleine, chevauche l'item
 * suivant - confirme par l'utilisateur comme absent du fichier PowerPoint
 * d'origine, ou l'espace entre les deux est simplement tres reduit ("quasiment
 * pas d'espace"), jamais un chevauchement reel.
 *
 * <p><b>Prise en charge retenue</b> : pour une forme identifiee par {@link
 * OverflowAwareTextFitter#isAutofitBroadeningExempt} (methode
 * package-private partagee entre les deux classes, meme paquetage - englobe
 * depuis le 2026-09-05 a la fois l'elargissement "sommaire" d'origine et
 * l'elargissement general experimental, voir sa Javadoc), le
 * garde-fou `textHeight > anchor.getHeight()` ci-dessus est leve - cette
 * forme est prise en charge par ce correctif MEME quand elle depasse sa
 * propre ancre, avec exactement la meme strategie a deux leviers que pour
 * toute autre forme (reduction d'interligne d'abord, repli sur la police
 * seulement si necessaire) : coherent avec le retour utilisateur ("quasiment
 * pas d'espace" chez PowerPoint, jamais un veritable manque de place), qui
 * suggere un leger exces d'interligne mesure plutot qu'un besoin de reduire
 * la police. Pour toute autre forme depassant sa propre ancre (le cas
 * general, non exemptee), le garde-fou reste inchange : {@link
 * OverflowAwareTextFitter} continue de s'en charger seul, sans aucune
 * interference de ce correctif.
 *
 * <p><b>Limite assumee (perimetre de collision)</b> : comme {@link
 * OverflowAwareTextFitter#findCollidingShape}, la comparaison se fait contre
 * l'ancre ENTIERE de la forme voisine (pas seulement la portion de cette
 * ancre reellement occupee par son propre texte) - une simplification deja
 * acceptee ailleurs dans ce projet.
 *
 * <p><b>Limite assumee (portee de la correction, toujours la forme entiere)</b> :
 * une piste plus chirurgicale a ete envisagee (ne reduire que les paragraphes
 * PRECEDANT le paragraphe en collision, jamais celui-ci ni ceux qui le
 * suivent) mais ecartee : elle romprait la compatibilite avec le cas d'origine
 * (slide 10), ou le chevauchement se resout en reduisant l'ensemble du bloc
 * et non un prefixe restreint. Consequence acceptee : si le repli sur la
 * police est necessaire, des paragraphes visibles SANS RAPPORT avec le
 * paragraphe reellement en cause (ex. les paragraphes 1, 3 et 5 du slide 17,
 * avant le paragraphe fautif) peuvent eux aussi voir leur police legerement
 * reduite - un effet de bord mineur et borne (jamais en dessous de {@link
 * #FONT_MIN_SCALE}), du meme ordre que celui deja accepte pour le cas
 * d'origine.
 *
 * <p><b>Limite assumee (cout)</b> : la mesure par paragraphe ajoute, pour
 * chaque forme ou une collision "forme entiere" est detectee, la creation
 * d'une copie de mesure jetable et plusieurs dizaines d'appels
 * supplementaires a {@code getTextHeight()} sur cette copie (un par
 * paragraphe teste, a chaque iteration de reduction). Attenue par le filtre
 * rapide existant (le test "forme entiere" bon marche, deja en place, ecarte
 * la grande majorite des formes d'un slide typique avant meme de construire
 * une copie de mesure) - reste a surveiller sur des formes a tres nombreux
 * paragraphes.
 *
 * <h2>Correction (2026-08-28, suite au retour de {@code mvn verify}) : recalage
 * sur la hauteur reelle plutot que confiance aveugle dans la copie de mesure</h2>
 * <p>La toute premiere version de cette 2e variante calculait la position de
 * CHAQUE paragraphe uniquement a partir des hauteurs cumulees rendues par
 * {@link ParagraphMeasurer} (la copie de mesure), y compris pour la hauteur
 * TOTALE du bloc - abandonnant au passage la pratique, jusque-la constante
 * dans cette classe (et dans {@link OverflowAwareTextFitter}), de recalculer
 * {@code ts.getTextHeight(graphics)} sur la forme REELLE apres chaque
 * ajustement d'interligne/police. Consequence observee lors du retour utilisateur
 * (test {@code fixNeighborOverlaps_resolvesCollision_whenVisibleParagraphAfterBlankSpacersOverlaps},
 * echec {@code expected: <1> but was: <0>}) : la copie de mesure, bien que
 * construite pour reproduire fidelement tout ce qui est cense influencer
 * {@code getTextHeight()} (largeur, marges, interligne, texte/taille/police de
 * chaque run, run fantome pour les paragraphes vides), en pratique avec de
 * vraies polices systeme ne reproduit pas la hauteur totale reelle au bit pres
 * (candidats plausibles : {@code getSpaceBefore()}/{@code getSpaceAfter()},
 * jusqu'ici non recopies - voir correction ci-dessous - ou des ecarts plus
 * subtils de {@code TextLayout} sur une forme neuve vs la forme source). Sur le
 * fichier de l'utilisateur, cet ecart d'estimation (systematiquement par
 * exces) empechait le test de collision par paragraphe de jamais constater la
 * disparition du chevauchement, meme une fois interligne et police reduits a
 * leurs limites - la correction s'abandonnait alors integralement (valeurs
 * d'origine restaurees, {@code count} inchange), au lieu de s'appliquer.
 *
 * <p><b>Correction retenue</b> : ne plus jamais traiter la hauteur TOTALE
 * rendue par la copie de mesure comme une verite absolue - seule
 * {@code ts.getTextHeight(graphics)}, recalculee sur la forme REELLE apres
 * chaque ajustement (exactement comme le faisait deja la version "forme
 * entiere" d'origine, et comme le fait toujours {@link
 * OverflowAwareTextFitter}), fait foi pour la hauteur totale et pour le calcul
 * de {@code baseY} (utile aux alignements {@code BOTTOM}/{@code MIDDLE}). La
 * copie de mesure ne sert plus qu'a etablir la REPARTITION RELATIVE de cette
 * hauteur totale, deja connue et fiable, entre les paragraphes : chaque
 * hauteur cumulee {@code measurer.cumulativeHeight(i)} est mise a l'echelle
 * par le facteur {@code ts.getTextHeight(graphics) / measurer.cumulativeHeight(dernierIndex)}
 * avant d'etre utilisee. Ce recalage rend le decoupage par paragraphe
 * insensible a un biais systematique (dans un sens ou dans l'autre) de la
 * copie de mesure - seule sa capacite a estimer des PROPORTIONS entre
 * paragraphes reste necessaire, une exigence bien plus faible que l'exactitude
 * absolue. Il preserve aussi exactement le raisonnement de non-regression
 * deja documente plus haut : pour une forme ou le DERNIER paragraphe est
 * visible (cas de la forme d'origine, slide 10), le rectangle de ce dernier
 * paragraphe se termine exactement a {@code baseY + ts.getTextHeight(graphics)}
 * - rigoureusement le meme resultat que l'ancien test "forme entiere".
 *
 * <p><b>Correction complementaire</b> : {@link #copyParagraphStructure} ne
 * recopiait pas {@code getSpaceBefore()}/{@code getSpaceAfter()} sur la copie
 * de mesure (seul l'interligne l'etait) - un ecart reel, corrige au passage
 * par prudence, meme si le recalage ci-dessus rend deja la mesure globalement
 * insensible a ce genre d'ecart residuel.
 *
 * <h2>Exemption des bulles narratives (callouts) de tout chevauchement (2026-09-03)</h2>
 * <p>Voir Javadoc de {@link OverflowAwareTextFitter}, meme section, pour le
 * recit complet (regression slide 4, "Title 3" retreci a 24% a cause d'un
 * "chevauchement" avec "Bulle narrative : rectangle a coins arrondis 1" qui
 * n'existe pas visuellement dans PowerPoint - artefact de la limite "Limite
 * assumee (perimetre de collision)" ci-dessus, l'ancre d'une bulle de rappel
 * englobant son bec). Toute forme identifiee par {@link
 * OverflowAwareTextFitter#isCalloutShape} (partagee, package-private) est
 * desormais : (1) ignoree en tout debut de {@link #fixNeighborOverlaps} - une
 * bulle n'est plus jamais elle-meme corrigee a cause d'une collision avec une
 * forme voisine ; (2) exclue de {@link #findCollidingShape} - une bulle n'est
 * plus jamais retenue comme forme "colliding" faisant chevaucher une AUTRE
 * forme. Ces bulles sont, dans la composition d'origine, deliberement
 * replacees au premier plan (ordre Z) a la fin, pour pointer par-dessus
 * d'autres formes : un chevauchement avec elles est donc toujours intentionnel.
 */
public final class NeighborShapeOverlapFixer {

    private static final Logger LOG = LoggerFactory.getLogger(NeighborShapeOverlapFixer.class);

    private static final double SPACING_STEP = 0.02;
    private static final double SPACING_MIN_FACTOR = 0.85;
    private static final int SPACING_MAX_ITER = 10;

    private static final double FONT_STEP = 0.02;
    private static final double FONT_MIN_SCALE = 0.25;
    private static final int FONT_MAX_ITER = 38;

    private NeighborShapeOverlapFixer() {
    }

    /**
     * Corrige, in-place, les formes de texte du slide dont le contenu
     * reellement mesure - sans jamais depasser leur propre ancre - chevauche
     * neanmoins une autre forme de texte independante, PARAGRAPHE PAR
     * PARAGRAPHE (voir Javadoc de la classe). A appeler apres {@link
     * OverflowAwareTextFitter#fitOverflowingText}, avant {@code
     * slide.draw(graphics)}.
     *
     * @return le nombre de formes effectivement corrigees (interligne et/ou police).
     */
    public static int fixNeighborOverlaps(XSLFSlide slide, Graphics2D graphics) {
        return fixNeighborOverlaps(slide, graphics, false);
    }

    /**
     * Variante de {@link #fixNeighborOverlaps(XSLFSlide, Graphics2D)} exposant
     * l'elargissement EXPERIMENTAL partage avec {@link
     * OverflowAwareTextFitter#isAutofitBroadeningExempt} - voir la Javadoc de
     * cette derniere classe, section "Elargissement general (experimental)".
     * A appeler avec la MEME valeur de {@code broadenAutofitExemption} que
     * celle passee a {@link OverflowAwareTextFitter#fitOverflowingText(
     * XSLFSlide, Graphics2D, boolean)} pour le meme slide (voir {@code
     * PptxSlideRenderer}, qui les invoque l'une apres l'autre avec la meme
     * valeur, issue de {@code RenderOptions.isBroadenAutofitExemption()}).
     *
     * @return le nombre de formes effectivement corrigees (interligne et/ou police).
     */
    public static int fixNeighborOverlaps(XSLFSlide slide, Graphics2D graphics, boolean broadenAutofitExemption) {
        int count = 0;
        List<XSLFShape> allTextShapes = collectTextShapes(slide.getShapes());
        // Voir Javadoc de la classe, section "Formes exemptees de l'elargissement
        // sommaire" : calcule une seule fois par slide, meme motif que dans
        // OverflowAwareTextFitter.fitOverflowingText.
        boolean sommaireSlide = OverflowAwareTextFitter.isSommaireSlide(slide);

        for (XSLFShape shape : allTextShapes) {
            XSLFTextShape ts = (XSLFTextShape) shape;
            // Voir Javadoc de la classe, section "Exemption des bulles narratives
            // (callouts) de tout chevauchement" : une bulle n'est jamais elle-meme
            // corrigee a cause d'une collision avec une forme voisine.
            if (OverflowAwareTextFitter.isCalloutShape(ts)) {
                continue;
            }
            Rectangle2D anchor = ts.getAnchor();
            if (anchor == null || anchor.getHeight() <= 0) {
                continue;
            }

            double textHeight = ts.getTextHeight(graphics);
            // Voir Javadoc de la classe, section "Formes exemptees de l'elargissement
            // sommaire" : une forme que OverflowAwareTextFitter a deliberement exemptee
            // de tout retrecissement (y compris quand elle depasse sa propre ancre)
            // reste a la charge de CE correctif meme dans ce cas - pour toute autre
            // forme, le depassement de sa propre ancre reste entierement gere par
            // OverflowAwareTextFitter, sans interference ici.
            boolean sommaireBroadeningExempt =
                    OverflowAwareTextFitter.isAutofitBroadeningExempt(ts, sommaireSlide, broadenAutofitExemption);
            if (textHeight > anchor.getHeight() && !sommaireBroadeningExempt) {
                continue;
            }

            List<XSLFTextParagraph> paragraphs = ts.getTextParagraphs();
            if (paragraphs.isEmpty()) {
                continue;
            }

            VerticalAlignment valign = ts.getVerticalAlignment();

            // Filtre rapide bon marche (inchange depuis la version initiale) : si le bloc
            // ENTIER (espaceurs vides compris) ne chevauche rien nulle part, alors AUCUN
            // paragraphe individuel ne peut chevaucher quoi que ce soit non plus (sa zone
            // est toujours incluse dans celle du bloc entier) - evite de construire une
            // copie de mesure pour la grande majorite des formes d'un slide typique.
            Rectangle2D wholeOccupied = computeOccupiedZone(anchor, textHeight, valign);
            if (findCollidingShape(wholeOccupied, ts, allTextShapes) == null) {
                continue;
            }

            try (ParagraphMeasurer measurer = new ParagraphMeasurer(slide, ts, anchor, graphics)) {
                XSLFShape colliding = findParagraphLevelCollision(ts, paragraphs, anchor, valign, allTextShapes, measurer, textHeight);
                if (colliding == null) {
                    // Le chevauchement "forme entiere" detecte ci-dessus n'atteint en realite
                    // aucun paragraphe VISIBLE (motif des espaceurs vides places devant une
                    // annotation, voir Javadoc de classe, "2e variante") : rien a corriger.
                    continue;
                }

                if (LOG.isDebugEnabled()) {
                    LOG.debug("{} : au moins un paragraphe visible chevauche '{}' -> tentative de correction "
                                    + "(interligne puis, en repli, police)",
                            shape.getShapeName(), colliding.getShapeName());
                }

                Map<XSLFTextParagraph, Double> spacingBaseline = captureSpacingBaseline(ts);
                Map<XSLFTextRun, Double> fontBaseline = captureFontBaseline(ts);

                boolean spacingAdjusted = false;
                double spacingFactor = 1.0;
                if (!spacingBaseline.isEmpty()) {
                    int iter = 0;
                    while (colliding != null && spacingFactor > SPACING_MIN_FACTOR && iter < SPACING_MAX_ITER) {
                        spacingFactor = Math.max(SPACING_MIN_FACTOR, spacingFactor - SPACING_STEP);
                        for (Map.Entry<XSLFTextParagraph, Double> e : spacingBaseline.entrySet()) {
                            e.getKey().setLineSpacing(effectiveSpacing(e.getValue()) * spacingFactor);
                        }
                        textHeight = ts.getTextHeight(graphics);
                        colliding = findParagraphLevelCollision(ts, paragraphs, anchor, valign, allTextShapes, measurer, textHeight);
                        iter++;
                        spacingAdjusted = true;
                    }
                }

                boolean fontAdjusted = false;
                double fontFactor = 1.0;
                if (colliding != null && !fontBaseline.isEmpty()) {
                    int iter = 0;
                    while (colliding != null && fontFactor > FONT_MIN_SCALE && iter < FONT_MAX_ITER) {
                        fontFactor -= FONT_STEP;
                        for (Map.Entry<XSLFTextRun, Double> e : fontBaseline.entrySet()) {
                            e.getKey().setFontSize(Math.max(1.0, e.getValue() * fontFactor));
                        }
                        textHeight = ts.getTextHeight(graphics);
                        colliding = findParagraphLevelCollision(ts, paragraphs, anchor, valign, allTextShapes, measurer, textHeight);
                        iter++;
                        fontAdjusted = true;
                    }
                }

                if (colliding != null) {
                    // Ni l'interligne (jusqu'a sa limite conservatrice) ni la police (jusqu'a
                    // MIN_SCALE) n'ont suffi : on restaure tout plutot que de livrer un
                    // resultat partiel qui chevauche quand meme la forme voisine.
                    if (spacingAdjusted) {
                        for (Map.Entry<XSLFTextParagraph, Double> e : spacingBaseline.entrySet()) {
                            e.getKey().setLineSpacing(e.getValue());
                        }
                    }
                    if (fontAdjusted) {
                        for (Map.Entry<XSLFTextRun, Double> e : fontBaseline.entrySet()) {
                            e.getKey().setFontSize(e.getValue());
                        }
                    }
                    if (LOG.isDebugEnabled()) {
                        LOG.debug("{} : correction impossible (interligne et police epuises), chevauchement "
                                + "toujours present -> taille et interligne d'origine restaures", shape.getShapeName());
                    }
                    continue;
                }

                if (spacingAdjusted || fontAdjusted) {
                    count++;
                    if (LOG.isDebugEnabled()) {
                        LOG.debug("{} corrigee : interligne a {}%{} (au moins un paragraphe visible chevauchait "
                                        + "une forme voisine)",
                                shape.getShapeName(), Math.round(spacingFactor * 100),
                                fontAdjusted ? (", police a " + Math.round(fontFactor * 100) + "% (repli)") : "");
                    }
                }
            }
        }
        return count;
    }

    /** {@code spacing} tel que stocke (peut etre {@code null}, herite/par defaut) ramene a une valeur exploitable. */
    private static double effectiveSpacing(Double spacing) {
        return (spacing == null || spacing <= 0) ? 100.0 : spacing;
    }

    /**
     * Paragraphes eligibles a la reduction d'interligne (interligne en
     * pourcentage ou non declare - voir {@link OversizedWhitespaceRunFixer}
     * pour la meme convention) et leur valeur d'origine exacte (potentiellement
     * {@code null}), pour permettre une restauration fidele en cas d'echec.
     * Un interligne en points absolus ({@code spcPts}, valeur negative) est
     * exclu : sa hauteur ne depend pas des metriques de police, le reduire ne
     * corrigerait rien de la surestimation visee ici.
     */
    private static Map<XSLFTextParagraph, Double> captureSpacingBaseline(XSLFTextShape ts) {
        Map<XSLFTextParagraph, Double> baseline = new HashMap<>();
        for (XSLFTextParagraph para : ts.getTextParagraphs()) {
            Double spacing = para.getLineSpacing();
            if (spacing == null || spacing > 0) {
                baseline.put(para, spacing);
            }
        }
        return baseline;
    }

    private static Map<XSLFTextRun, Double> captureFontBaseline(XSLFTextShape ts) {
        Map<XSLFTextRun, Double> baseline = new HashMap<>();
        for (XSLFTextParagraph para : ts.getTextParagraphs()) {
            for (XSLFTextRun run : para.getTextRuns()) {
                Double size = run.getFontSize();
                if (size != null) {
                    baseline.put(run, size);
                }
            }
        }
        return baseline;
    }

    /**
     * Rectangle reellement occupe par le texte mesure d'une forme, selon son
     * alignement vertical resolu (TOP par defaut si non precise, convention
     * OOXML {@code anchor="t"}) - contrairement a {@link
     * OverflowAwareTextFitter#computeOverflowZones}, qui ne calcule que la
     * portion au-dela de l'ancre, ceci renvoie la zone occupee dans son
     * ensemble (qu'elle depasse ou non l'ancre - ici appele uniquement quand
     * elle ne la depasse pas, voir {@link #fixNeighborOverlaps}).
     */
    static Rectangle2D computeOccupiedZone(Rectangle2D anchor, double textHeight, VerticalAlignment valign) {
        VerticalAlignment v = (valign == null) ? VerticalAlignment.TOP : valign;
        double y;
        switch (v) {
            case BOTTOM:
                y = anchor.getY() + anchor.getHeight() - textHeight;
                break;
            case MIDDLE:
                y = anchor.getY() + (anchor.getHeight() - textHeight) / 2.0;
                break;
            case TOP:
            default:
                y = anchor.getY();
                break;
        }
        return new Rectangle2D.Double(anchor.getX(), y, anchor.getWidth(), Math.max(textHeight, 0));
    }

    /**
     * Comme {@link OverflowAwareTextFitter#findCollidingShape} : ignore les
     * formes sans texte visible (fonds/panneaux), et compare a l'ancre
     * ENTIERE de la forme voisine (voir "Limite assumee" en Javadoc de
     * classe).
     */
    static XSLFShape findCollidingShape(Rectangle2D occupied, XSLFTextShape self, List<XSLFShape> allTextShapes) {
        for (XSLFShape other : allTextShapes) {
            if (other == self || !(other instanceof XSLFTextShape) || OverflowAwareTextFitter.isCalloutShape(other)) {
                continue;
            }
            XSLFTextShape ots = (XSLFTextShape) other;
            String text = ots.getText();
            if (text == null || text.trim().isEmpty()) {
                continue;
            }
            Rectangle2D otherAnchor = ots.getAnchor();
            if (otherAnchor != null && occupied.intersects(otherAnchor)) {
                return other;
            }
        }
        return null;
    }

    /**
     * Parcourt les paragraphes de {@code ts} dans l'ordre et renvoie la
     * premiere forme voisine chevauchee par un paragraphe VISIBLE (texte non
     * entierement blanc) - les paragraphes entierement vides (espaceurs) sont
     * ignores en tant que source de collision, meme si leur zone geometrique
     * chevauche une forme voisine (voir Javadoc de classe, "2e variante").
     * {@code measurer} donne la REPARTITION RELATIVE de la hauteur totale
     * entre paragraphes - jamais la hauteur totale elle-meme, qui reste
     * {@code realTextHeight} ({@code ts.getTextHeight(graphics)}, recalculee
     * sur la forme REELLE par l'appelant avant chaque appel). Voir "Correction
     * (2026-08-28...)" en Javadoc de classe pour la justification de ce
     * recalage.
     */
    private static XSLFShape findParagraphLevelCollision(XSLFTextShape ts, List<XSLFTextParagraph> paragraphs,
            Rectangle2D anchor, VerticalAlignment valign, List<XSLFShape> allTextShapes, ParagraphMeasurer measurer,
            double realTextHeight) {
        double measuredTotal = measurer.cumulativeHeight(paragraphs.size() - 1);
        // Met a l'echelle la repartition mesuree par la copie sur la hauteur REELLE
        // et fiable de la forme source - rend le decoupage par paragraphe insensible
        // a un biais systematique (dans un sens ou dans l'autre) de la copie de mesure.
        double scale = (measuredTotal > 0) ? (realTextHeight / measuredTotal) : 1.0;
        double baseY = computeOccupiedZone(anchor, realTextHeight, valign).getY();

        double prevCumulative = 0;
        for (int i = 0; i < paragraphs.size(); i++) {
            double cumulative = measurer.cumulativeHeight(i) * scale;
            if (isVisibleParagraph(paragraphs.get(i))) {
                double paragraphHeight = Math.max(cumulative - prevCumulative, 0);
                Rectangle2D paragraphRect = new Rectangle2D.Double(
                        anchor.getX(), baseY + prevCumulative, anchor.getWidth(), paragraphHeight);
                XSLFShape colliding = findCollidingShape(paragraphRect, ts, allTextShapes);
                if (colliding != null) {
                    return colliding;
                }
            }
            prevCumulative = cumulative;
        }
        return null;
    }

    /** Un paragraphe est "visible" des qu'au moins un de ses runs contient du texte non entierement blanc. */
    private static boolean isVisibleParagraph(XSLFTextParagraph para) {
        for (XSLFTextRun run : para.getTextRuns()) {
            String text = run.getRawText();
            if (text != null && !text.trim().isEmpty()) {
                return true;
            }
        }
        return false;
    }

    /** Parcourt les formes du slide, y compris a l'interieur des groupes - meme perimetre que {@link OverflowAwareTextFitter}. */
    private static List<XSLFShape> collectTextShapes(List<XSLFShape> shapes) {
        List<XSLFShape> result = new ArrayList<>();
        for (XSLFShape shape : shapes) {
            if (shape instanceof XSLFGroupShape) {
                result.addAll(collectTextShapes(((XSLFGroupShape) shape).getShapes()));
            } else if (shape instanceof XSLFTextShape) {
                result.add(shape);
            }
        }
        return result;
    }

    /**
     * Mesure la position verticale cumulee de chaque paragraphe d'une forme
     * source, en l'absence de toute API POI publique exposant cette mesure
     * (voir Javadoc de classe). Construit une COPIE DE MESURE jetable ({@link
     * XSLFTextBox}, ajoutee au meme slide puis retiree par {@link #close()},
     * jamais dessinee) et la tronque/reconstruit progressivement - la forme
     * SOURCE, elle, n'est JAMAIS structurellement modifiee (ses proprietes de
     * police/interligne le sont, par {@link #fixNeighborOverlaps}, mais
     * jamais l'identite de ses paragraphes/runs).
     */
    private static final class ParagraphMeasurer implements AutoCloseable {

        private final XSLFSlide slide;
        private final XSLFTextBox scratch;
        private final Graphics2D graphics;
        private final List<XSLFTextParagraph> sourceParagraphs;

        ParagraphMeasurer(XSLFSlide slide, XSLFTextShape source, Rectangle2D anchor, Graphics2D graphics) {
            this.slide = slide;
            this.graphics = graphics;
            this.sourceParagraphs = source.getTextParagraphs();

            this.scratch = slide.createTextBox();
            double width = Math.max(anchor.getWidth(), 1.0);
            double generousHeight = Math.max(anchor.getHeight(), 1.0) * 8 + 10_000;
            scratch.setAnchor(new Rectangle2D.Double(0, 0, width, generousHeight));
            scratch.setWordWrap(source.getWordWrap());
            scratch.setTextAutofit(TextShape.TextAutofit.NONE);
            scratch.setLeftInset(source.getLeftInset());
            scratch.setRightInset(source.getRightInset());
            scratch.setTopInset(source.getTopInset());
            scratch.setBottomInset(source.getBottomInset());
            // Le textbox neuf demarre avec un paragraphe vide par defaut - retire tout de
            // suite, cumulativeHeight() construit exactement ce qu'il faut au fil des appels.
            scratch.removeTextParagraph(scratch.getTextParagraphs().get(0));
        }

        /** Hauteur cumulee mesuree des paragraphes [0, throughIndexInclusive] - 0.0 pour un index negatif. */
        double cumulativeHeight(int throughIndexInclusive) {
            if (throughIndexInclusive < 0) {
                return 0.0;
            }
            int desiredCount = throughIndexInclusive + 1;

            List<XSLFTextParagraph> current = scratch.getTextParagraphs();
            while (current.size() > desiredCount) {
                scratch.removeTextParagraph(current.get(current.size() - 1));
                current = scratch.getTextParagraphs();
            }
            while (current.size() < desiredCount) {
                int nextIndex = current.size();
                XSLFTextParagraph fresh = scratch.addNewTextParagraph();
                copyParagraphStructure(sourceParagraphs.get(nextIndex), fresh);
                current = scratch.getTextParagraphs();
            }
            // Les paragraphes deja presents peuvent avoir change (interligne/police en
            // cours d'essai sur la forme source depuis le dernier appel) - toujours
            // resynchroniser avant de mesurer.
            for (int i = 0; i < desiredCount; i++) {
                refreshParagraphValues(sourceParagraphs.get(i), current.get(i));
            }
            return scratch.getTextHeight(graphics);
        }

        @Override
        public void close() {
            slide.removeShape(scratch);
        }
    }

    /**
     * Reproduit sur {@code to} (un paragraphe de la copie de mesure, tout
     * juste ajoute) ce qui influence la hauteur mesuree par Java2D pour
     * {@code from} (un paragraphe de la forme source) : texte, taille et
     * police DEJA RESOLUE (voir Javadoc de classe) de chaque run, et
     * interligne du paragraphe. Un paragraphe SANS AUCUN run (espaceur pur,
     * {@code endParaRPr} seul) est reproduit par un run fantome contenant un
     * espace a {@link XSLFTextParagraph#getDefaultFontSize()} - le mecanisme
     * que POI utilise lui-meme en interne pour ne jamais mesurer une hauteur
     * nulle sur un tel paragraphe (voir Javadoc de classe).
     */
    private static void copyParagraphStructure(XSLFTextParagraph from, XSLFTextParagraph to) {
        List<XSLFTextRun> fromRuns = from.getTextRuns();
        if (fromRuns.isEmpty()) {
            XSLFTextRun phantom = to.addNewTextRun();
            phantom.setText(" ");
            Double defaultSize = from.getDefaultFontSize();
            if (defaultSize != null) {
                phantom.setFontSize(defaultSize);
            }
        } else {
            for (XSLFTextRun fromRun : fromRuns) {
                XSLFTextRun toRun = to.addNewTextRun();
                String text = fromRun.getRawText();
                toRun.setText(text == null ? "" : text);
                Double size = fromRun.getFontSize();
                if (size != null) {
                    toRun.setFontSize(size);
                }
                String family = fromRun.getFontFamily();
                if (family != null) {
                    toRun.setFontFamily(family);
                }
            }
        }
        to.setLineSpacing(from.getLineSpacing());
        to.setSpaceBefore(from.getSpaceBefore());
        to.setSpaceAfter(from.getSpaceAfter());
    }

    /**
     * Resynchronise ce que {@link #fixNeighborOverlaps} peut avoir change
     * depuis le dernier appel (interligne du paragraphe, taille des runs
     * REELS) - jamais le texte, la police ou la structure (runs
     * ajoutes/retires), que ce correctif ne modifie jamais. {@code
     * spaceBefore}/{@code spaceAfter} sont aussi resynchronises par
     * simplicite/coherence avec {@link #copyParagraphStructure}, bien que ce
     * correctif ne les modifie jamais lui-meme. Sans effet sur la taille pour
     * un paragraphe-espaceur (son run fantome n'est lie a aucune propriete de
     * run modifiee par ce correctif).
     */
    private static void refreshParagraphValues(XSLFTextParagraph from, XSLFTextParagraph to) {
        to.setLineSpacing(from.getLineSpacing());
        to.setSpaceBefore(from.getSpaceBefore());
        to.setSpaceAfter(from.getSpaceAfter());
        List<XSLFTextRun> fromRuns = from.getTextRuns();
        if (fromRuns.isEmpty()) {
            return;
        }
        List<XSLFTextRun> toRuns = to.getTextRuns();
        for (int i = 0; i < fromRuns.size() && i < toRuns.size(); i++) {
            Double size = fromRuns.get(i).getFontSize();
            if (size != null) {
                toRuns.get(i).setFontSize(size);
            }
        }
    }
}
