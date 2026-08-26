package io.github.atlan77c.pptx2picture.internal;

import org.apache.poi.sl.draw.DrawConnectorShape;
import org.apache.poi.sl.draw.DrawFactory;
import org.apache.poi.sl.draw.Drawable;
import org.apache.poi.sl.draw.geom.Outline;
import org.apache.poi.sl.draw.geom.Path;
import org.apache.poi.sl.usermodel.ConnectorShape;
import org.apache.poi.sl.usermodel.LineDecoration;
import org.apache.poi.sl.usermodel.LineDecoration.DecorationShape;
import org.apache.poi.sl.usermodel.LineDecoration.DecorationSize;
import org.apache.poi.sl.usermodel.PaintStyle.PaintModifier;
import org.apache.poi.sl.usermodel.ShapeType;
import org.apache.poi.xslf.usermodel.XSLFConnectorShape;
import org.openxmlformats.schemas.drawingml.x2006.main.CTShapeProperties;
import org.openxmlformats.schemas.presentationml.x2006.main.CTConnector;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.awt.BasicStroke;
import java.awt.Graphics2D;
import java.awt.Paint;
import java.awt.Shape;
import java.awt.Stroke;
import java.awt.Color;
import java.awt.geom.AffineTransform;
import java.awt.geom.Area;
import java.awt.geom.Ellipse2D;
import java.awt.geom.Path2D;
import java.awt.geom.PathIterator;
import java.awt.geom.Rectangle2D;
import java.util.Collection;
import java.util.Locale;

/**
 * Corrige un ecart de fidelite decouvert lors d'une revue visuelle systematique
 * du rendu : les pointes de fleche des connecteurs courbes
 * ({@code curvedConnector2/3/4/5}) et coudes ({@code bentConnector2/3/4/5})
 * sont mal orientees, alors que le trait du connecteur lui-meme se dessine
 * correctement.
 *
 * <p><b>Cause racine, identifiee dans le code source d'Apache POI lui-meme</b>
 * ({@code DrawSimpleShape.getHeadDecoration()}/{@code getTailDecoration()}) :
 * l'angle de la pointe de fleche n'est jamais calcule a partir de la tangente
 * reelle du trace au point d'arrivee, seulement a partir du rapport
 * hauteur/largeur de la boite englobante de la forme :
 * <pre>{@code
 * alpha = Math.atan(anchor.getHeight() / anchor.getWidth());
 * }</pre>
 * Ce calcul suppose implicitement que le connecteur EST la diagonale de sa
 * boite englobante - exact pour {@code straightConnector1} (PowerPoint encode
 * alors les autres orientations via {@code flipH}/{@code flipV}), mais faux
 * pour un connecteur courbe ou coude, dont le trace entre et sort de chaque
 * extremite a l'horizontale ou a la verticale, jamais en diagonale. Confirme
 * par mesure directe sur des rendus reels : les connecteurs en cause sont bien
 * {@code curvedConnector3}, {@code bentConnector2} ou {@code bentConnector3}
 * (jamais {@code straightConnector1}) ; la rotation ({@code rot} sur le
 * {@code xfrm}) presente sur certains d'entre eux n'est pas la cause (des
 * instances sans aucune rotation presentent le meme defaut) - juste un facteur
 * independant qui peut aggraver visuellement l'erreur deja presente sans elle.
 *
 * <p><b>[v13-drawfactory-zorder] Architecture retenue - substitution via
 * {@code DrawFactory}, plus de retrait/redessin</b>. Les versions v9 a v12 de ce
 * correctif retiraient le connecteur de la slide ({@code slide.removeShape()})
 * avant {@code slide.draw(graphics)}, puis le redessinaient nous-memes APRES -
 * ce qui a resolu les problemes d'orientation et de recouvrement pres de la
 * pointe (voir plus bas), mais a introduit une regression decouverte lors
 * d'une revue visuelle ulterieure : le connecteur redessine se retrouvait
 * <b>systematiquement au premier plan</b>, au-dessus de toutes les autres formes
 * de la slide - y compris celles sous lesquelles il aurait du normalement
 * passer - puisqu'il n'etait plus dessine a sa place naturelle dans l'ordre
 * d'empilement (z-order) des formes, mais toujours en tout dernier.
 *
 * <p>La solution retenue s'appuie sur un point d'extension officiel et
 * documente d'Apache POI (verifie directement dans le code source, module
 * {@code org.apache.poi.sl.draw}, version 5.2.5) : {@code slide.draw(graphics)}
 * delegue le dessin de chaque forme (y compris recursivement dans les groupes)
 * a un {@code Drawable} obtenu via {@code DrawFactory.getInstance(graphics)
 * .getDrawable(shape)}, et ce {@code DrawFactory} peut etre remplace par une
 * sous-classe personnalisee, installee sur le {@code Graphics2D} via le hint de
 * rendu {@code Drawable.DRAW_FACTORY} - voir {@link #installBeforeDraw}. Notre
 * {@code ArrowCorrectingDrawFactory} intercepte uniquement
 * {@code getDrawable(ConnectorShape)} pour les connecteurs courbes/coudes a
 * pointe declaree (voir {@link #qualifiesForCorrection}) et leur substitue
 * {@link TangentCorrectedConnector} ; pour tout le reste (formes normales, et
 * connecteurs qui n'ont pas besoin d'etre corriges), le comportement standard
 * de POI est utilise sans aucune modification ({@code super.getDrawable(...)}).
 *
 * <p>Consequence directe : {@code slide.draw()} continue d'appeler notre
 * dessinateur exactement a la place du connecteur dans la liste des formes de
 * la slide, avec le meme rituel que pour n'importe quelle autre forme (POI
 * applique lui-meme la transformation de la forme avant d'appeler
 * {@code draw()} - voir {@code DrawSheet}/{@code DrawShape.applyTransform}) :
 * <b>l'ordre d'empilement (z-order) redevient donc naturellement correct</b>,
 * sans aucun traitement special de notre part. Le connecteur n'est plus jamais
 * retire de la slide, ce qui elimine du meme coup, entierement et
 * definitivement, le risque {@code XmlValueDisconnected} qui avait motive
 * l'essentiel de la complexite des versions v11/v12 (capture-avant-retrait,
 * {@code PendingArrow}, redessin differe en espace device) - toute cette
 * mecanique disparait : {@link TangentCorrectedConnector#draw} dessine
 * directement, dans le repere local deja mis en place par POI au moment de
 * l'appel, sans avoir besoin de capturer ni de projeter quoi que ce soit vers
 * un espace "device" pour un usage differe.
 *
 * <p><b>Bonus non verifie</b> : puisque le hint {@code Drawable.DRAW_FACTORY}
 * est porte par le {@code Graphics2D} lui-meme, il reste actif y compris
 * lorsque POI descend recursivement dans un groupe de formes ({@code
 * DrawGroupShape} utilise le meme mecanisme) - un connecteur a l'interieur
 * d'un groupe devrait donc, en principe, egalement etre corrige desormais
 * (limite qui existait explicitement dans les versions v9 a v12). Non teste
 * sur un cas reel a ce jour : aucun connecteur a corriger rencontre jusqu'ici
 * n'etait a l'interieur d'un groupe.
 *
 * <p><b>Cas volontairement non gere</b> : un connecteur dont le trace lui-meme
 * comporte un {@code Outline} REELLEMENT rempli ({@code Path.isFilled()} - cas
 * jamais rencontre en pratique, un connecteur PowerPoint n'est normalement pas
 * remplissable depuis l'interface utilisateur) retombe entierement sur le
 * rendu standard de POI, bug d'orientation de pointe potentiellement inclus -
 * voir {@link TangentCorrectedConnector#draw}.
 * <p><b>[v14-detection-remplissage-corrigee]</b> Attention, {@code
 * getFillPaint(graphics)} seul n'est PAS ce signal : pour un connecteur tout a
 * fait normal, il renvoie presque toujours un {@code Paint} non-null (valeur
 * heritee/par defaut du theme, sans rapport avec un remplissage visible) - le
 * confondre avec "ce connecteur a un remplissage" a fait retomber, en v13, la
 * quasi-totalite des connecteurs a corriger sur le rendu standard de POI (donc
 * sur le bug d'orientation d'origine) - regression constatee sur un rendu reel
 * juste apres la livraison de v13, corrigee en v14. Motif de la
 * limitation elle-meme (inchangee) : la logique de remplissage de
 * {@code DrawSimpleShape.draw()} (calcul des degrades, {@code PaintModifier}...)
 * repose sur des methodes {@code private}
 * de POI, donc non reutilisables depuis une sous-classe ; la redevelopper pour
 * un cas jamais observe n'aurait apporte que du risque.
 *
 * <p><b>Angle de la pointe - trois approches essayees, dans cet ordre</b> (la
 * geometrie elle-meme n'a pas change entre v9 et v13, seul le moment/la
 * maniere dont elle est appliquee a change - voir ci-dessus) :
 * <ol>
 *   <li>la secante entre les deux derniers points d'un {@code PathIterator}
 *       <i>aplati</i> - ecartee, imprecise sur les courbes serrees ;</li>
 *   <li>la tangente Bezier <i>instantanee</i> au point de controle voisin
 *       de l'extremite (deriv exacte a t=1) - ecartee malgre son exactitude
 *       mathematique : pour {@code curvedConnector2/3} a adj1=50% (valeur
 *       par defaut), le dernier point de controle du dernier segment
 *       partage exactement la meme coordonnee "b" (bord bas/droit) que le
 *       point d'arrivee, donc la tangente y est rigoureusement horizontale
 *       ou verticale - un artefact du preset (pensé pour que le trait
 *       arrive perpendiculaire au bord de la boite cible), pas l'axe que
 *       l'oeil percoit sur le trace visible (mesure sur un rendu reel :
 *       ~45-60[deg]) ;</li>
 *   <li><b>retenue</b> : la corde d'une portion du segment proche de
 *       l'extremite (pas le segment entier, pas le point de controle
 *       seul) - voir {@link #FRACTION_FROM_TIP_SINGLE_SEGMENT} et la
 *       Javadoc de {@link TangentCorrectedConnector#computeEndpoint}.</li>
 * </ol>
 * Une etape intermediaire (corde du segment <i>entier</i>, point de depart
 * -&gt; point d'arrivee) avait ete essayee entre 2 et 3 : elle se reduit
 * exactement a l'ancien calcul {@code atan(h/w)} de POI pour un
 * {@code curvedConnector2/3} a ajustement par defaut (coincidence
 * mathematique : avec adj1=50%, la corde du dernier segment est alignee
 * sur la diagonale de la boite englobante), et donnait donc un resultat
 * plausible visuellement, mais une verification pixel sur un rendu reel
 * (mesure de la tangente locale de la courbe pres de son extremite,
 * comparee a l'axe d'un chevron de diagnostic) a montre un ecart residuel
 * de ~15-20[deg] sur des connecteurs utilisant {@code curvedConnector3},
 * la corde du segment entier "moyennant" sur toute sa longueur alors que le
 * segment est fortement courbe pres de l'extremite. D'ou le passage a une
 * corde locale, proche de la pointe seulement.
 *
 * <p><b>Beneice secondaire</b> : la position meme de la pointe (pas seulement son
 * angle) est corrigee au passage - le code d'origine de POI la place au coin de
 * la boite englobante ({@code anchor.getX()+anchor.getWidth()}, etc.), qui ne
 * coincide pas forcement avec l'extremite reelle du trace pour un connecteur
 * courbe/coude ; notre version la place au dernier point du trace reel. Verifie
 * par mesure pixel sur plusieurs connecteurs cibles (ecart pointe/trait
 * &lt; 4px dans tous les cas mesures).
 *
 * <p><b>Correctif complementaire - trait raccourci pres de la pointe</b> : sur
 * un connecteur avec une decoration pleine ({@code TRIANGLE}), un morceau de
 * trait de la meme couleur pouvait sembler depasser sur les cotes de la pointe
 * repeinte, sur une longueur non negligeable (jusqu'a environ un tiers de la
 * longueur de la pointe, mesure sur un rendu reel). Cause : le trait "brut" a
 * une largeur CONSTANTE sur toute sa longueur, alors qu'un triangle de pointe
 * s'amincit lineairement jusqu'a une largeur NULLE exactement au point
 * d'arrivee - sur la portion terminale du triangle (dont la largeur locale est
 * inferieure a celle du trait), le trait depasse donc visiblement de part et
 * d'autre si les deux sont peints l'un sur l'autre a pleine largeur.
 * <p>Plusieurs corrections essayees, dans cet ordre :
 * <ol>
 *   <li>deplacer le point d'ancrage de la pointe - ecartee : pour combler un
 *       ecart de cette ampleur (~1/3 de la longueur de la pointe), il aurait
 *       fallu decaler l'ancre au point de chevaucher la forme cible ;</li>
 *   <li>ajouter, uniquement pour {@code TRIANGLE}, un petit rectangle
 *       (largeur = celle reelle du trait) fusionne avec le triangle via
 *       {@code Area#add}, recouvrant la moitie de la longueur de la pointe la
 *       plus proche de son extremite - ecartee apres verification visuelle sur
 *       un rendu reel : meme de la bonne couleur, la fusion cree un "step"/une
 *       marche geometrique visible dans la silhouette (le rectangle a largeur
 *       constante rencontrant abruptement le triangle qui s'amincit) - un
 *       probleme de silhouette, pas de couleur, donc insoluble en changeant
 *       simplement la couleur du rectangle ;</li>
 *   <li>exclure cette meme zone (moitie de la longueur de la pointe la plus
 *       proche de son extremite) du clip du {@code Graphics2D}, juste avant
 *       {@code slide.draw(graphics)} (clip restaure juste apres) - ecartee a
 *       son tour apres verification visuelle sur un rendu reel : le clip d'un
 *       {@code Graphics2D} s'applique a TOUT ce qui est peint pendant qu'il est
 *       actif, pas seulement au connecteur vise - {@code slide.draw()} peint la
 *       slide ENTIERE en un seul appel, donc la meme exclusion rognait aussi la
 *       forme CIBLE pointee par la fleche, la ou son aire chevauche la petite
 *       zone exclue (ce qui est tres probable par construction : une fleche
 *       touche presque toujours sa cible pres de sa pointe) - visible comme une
 *       petite encoche blanche dans le bord de la forme cible, a l'endroit
 *       precis ou le trait aurait du s'arreter ;</li>
 *   <li>(v9-v12) retirer le connecteur de la slide et le redessiner nous-memes
 *       en entier apres {@code slide.draw()}, avec la meme exclusion
 *       geometrique que l'etape precedente mais appliquee UNIQUEMENT autour de
 *       notre propre appel a {@code graphics.draw(...)} pour le trait de CE
 *       connecteur : fonctionnait pour ce probleme precis, mais cause la
 *       regression de z-order documentee plus haut - superseded par v13.</li>
 * </ol>
 * <p><b>Retenue (v13)</b> : la meme exclusion geometrique (pres de la pointe,
 * uniquement pour {@code TRIANGLE}) qu'aux etapes precedentes, mais calculee et
 * appliquee <b>dans le repere local</b> du connecteur, directement au sein de
 * {@link TangentCorrectedConnector#draw} - donc uniquement autour du trait de
 * CE connecteur, jamais autour d'une autre forme (le clip est restaure avant
 * de dessiner la pointe elle-meme, puis restaure a nouveau avant de rendre la
 * main a POI) - voir {@link TangentCorrectedConnector#computeShaftExclusion}.
 * Le triangle repeint reste un triangle simple, sans rectangle fusionne - voir
 * {@link TangentCorrectedConnector#buildDecoration}, cas {@code TRIANGLE}.
 *
 * <p><b>[v15-sommets-arrondis] Sommets du triangle de pointe arrondis au lieu
 * d'anguleux</b> (ecart releve par comparaison directe avec le rendu PowerPoint
 * natif). Cause identifiee dans le code source d'Apache POI ({@code org.apache.poi.sl.draw
 * .geom.Path}, verifie sur la version 5.2.5) : un {@code Path} nouvellement
 * construit a {@code stroke = true} par defaut (donc {@code isStroked()} vaut
 * {@code true} tant que {@code setStroke(false)} n'est pas appele explicitement) -
 * seuls les cas {@code STEALTH}/{@code ARROW} de {@link TangentCorrectedConnector
 * #buildDecoration} configuraient explicitement leur {@code Path} (a dessein,
 * puisque ce sont des chevrons non remplis, dont le trace EST le contour), les cas
 * {@code OVAL}/{@code TRIANGLE} en heritaient donc silencieusement sans jamais le
 * desactiver. Consequence dans {@link TangentCorrectedConnector#drawCorrectedDecoration}
 * : la pointe {@code TRIANGLE} etait remplie ({@code isFilled()==true}, voulu),
 * PUIS son contour retrace PAR-DESSUS avec le {@code stroke} du connecteur
 * (jointures/extremites arrondies) - repasse invisible sur une grande pointe (la
 * largeur du trait y est negligeable), mais qui arrondit visiblement les 3 sommets
 * d'une petite pointe, la largeur de trait fixe n'etant plus negligeable devant sa
 * taille (effet proportionnel a la taille de la pointe, jamais a une valeur
 * absolue). Reproduit isolement (java.awt seul, hors POI, hors ambiguite de scale
 * ou d'antialiasing) : a largeur de trait fixe, une pointe de 12px ressort
 * nettement arrondie sur ses 3 sommets, une pointe de 100px reste anguleuse -
 * comportement identique observe sur un rendu reel (petites pointes arrondies,
 * grandes pointes nettes). Corrige en ajoutant {@code p.setStroke(false)} pour
 * {@code OVAL} et {@code TRIANGLE} : une pointe pleine n'a besoin que de son
 * remplissage, pas d'un contour trace en plus (contrairement a {@code ARROW}/
 * {@code STEALTH}) - voir {@link TangentCorrectedConnector#buildDecoration}.
 *
 * <p><b>[v16-stealth-polygone] Pointe {@code STEALTH} ("flechette" a barbes)
 * aux sommets emousses au lieu d'anguleux</b>. Meme famille de symptome que
 * [v15-sommets-arrondis] (comparaison directe avec une capture PowerPoint :
 * sommets nets d'un cote, emousses de l'autre - apex, barbes ET encoche
 * concave), mais cause differente et plus structurelle : contrairement a
 * {@code TRIANGLE}, {@code STEALTH} n'etait pas une forme pleine mal doublee
 * d'un trait superflu, mais un simple chevron a 2 segments trace au trait
 * (meme construction que {@code ARROW}, partagee avec le meme {@code case}
 * avant ce correctif) - une approche fidele a la construction d'origine de
 * POI, mais incapable par nature de produire des sommets nets : le sommet
 * central depend de la jointure du trait (rond/biseaute selon le style), et
 * les deux extremites OUVERTES du chevron (qui devraient former les barbes
 * pointues et l'encoche concave rentrante que montre PowerPoint) ne sont
 * jamais que des bouts de trait selon le "cap" configure - jamais un vrai
 * angle. Corrige en construisant {@code STEALTH} comme un veritable polygone
 * plein a 4 sommets (pointe, barbe, encoche concave, barbe - voir
 * {@link #STEALTH_NOTCH_FRACTION} et {@link TangentCorrectedConnector
 * #buildDecoration}, cas {@code STEALTH}), avec {@code setStroke(false)}
 * (lecon directe de [v15-sommets-arrondis] : un {@code Path()} par defaut
 * retrace sinon son propre contour par-dessus le remplissage, emoussant a
 * nouveau tous les sommets). {@code ARROW} reste inchange (un vrai chevron non
 * rempli trace au trait, y compris dans le rendu PowerPoint natif - pas le
 * meme defaut). Beneice secondaire : {@code STEALTH} devenant une forme pleine
 * qui s'amincit jusqu'a une largeur nulle a la pointe (comme {@code TRIANGLE}),
 * elle herite du meme risque de trait brut debordant pres de la pointe - voir
 * {@link TangentCorrectedConnector#computeShaftExclusion}, desormais applique
 * aux deux formes.
 */
public final class ConnectorArrowFixer {

    private static final Logger LOG = LoggerFactory.getLogger(ConnectorArrowFixer.class);

    /** Meme constante que {@code DrawSimpleShape.DECO_SIZE_POW} (non exposee publiquement par POI). */
    private static final double DECO_SIZE_POW = 1.5d;

    /**
     * [v5-fraction-adaptative] Fraction de la longueur (en parametre
     * {@code t}, pas en longueur d'arc) du segment terminal, mesuree depuis la
     * pointe, utilisee pour batir la corde qui sert d'axe a la pointe de
     * fleche. 0 donnerait la tangente instantanee exacte au point de controle
     * (approche rejetee, degenere a l'horizontale/verticale sur
     * {@code curvedConnector2/3}) ; 1 donnerait la corde du segment entier
     * (approche [v2-corde], encore ~15-20[deg] d'ecart mesure sur
     * {@code curvedConnector3}).
     *
     * <p>Une premiere valeur unique (0.15, [v4-souscorde]) a ete mesuree
     * correcte sur un rendu reel pour {@code curvedConnector2} (une seule
     * commande {@code cubicTo} dans le trace), mais encore ~15-20[deg] en
     * dessous de l'angle voulu pour {@code curvedConnector3} (deux commandes
     * {@code cubicTo}) - confirme par le balayage {@link #DEBUG_FRACTION_SWEEP}
     * journalise en DEBUG : plusieurs connecteurs {@code curvedConnector3}
     * mesures independamment convergent vers ~38-41% (pas ~15%) pour atteindre
     * l'angle voulu. Plausible : le deuxieme segment de {@code curvedConnector3}
     * porte une inflexion (forme en S) que {@code curvedConnector2} (un seul
     * cubicTo, sans inflexion) n'a pas ; une fraction plus grande, donc moins
     * locale, semble necessaire pour "sortir" de cette inflexion et retrouver
     * l'axe visuellement percu.
     * D'ou la distinction par nombre de segments du trace plutot qu'une
     * fraction unique : voir {@link #FRACTION_FROM_TIP_SINGLE_SEGMENT} et
     * {@link #FRACTION_FROM_TIP_MULTI_SEGMENT}, choisie dans
     * {@link TangentCorrectedConnector#tangentNearTip} selon le nombre total
     * de segments de l'{@code Outline} (compte dans
     * {@link TangentCorrectedConnector#computeEndpoint}).
     * {@link TangentCorrectedConnector#computeEndpoint} continue de
     * journaliser en DEBUG l'angle obtenu pour tout un balayage de valeurs
     * ({@link #DEBUG_FRACTION_SWEEP}) afin de pouvoir corriger ces constantes
     * sans nouveau cycle de mesure a l'aveugle si necessaire (pour un
     * {@code bentConnector} par exemple, pas encore mesure).
     */
    private static final double FRACTION_FROM_TIP_SINGLE_SEGMENT = 0.15d;

    /** Voir {@link #FRACTION_FROM_TIP_SINGLE_SEGMENT}. Valeur pour un trace a
     * 2 segments ou plus (ex. {@code curvedConnector3}, deux {@code cubicTo}) :
     * ~0.40, estimee par interpolation entre les points t=30% et t=50% du
     * {@link #DEBUG_FRACTION_SWEEP} mesure sur plusieurs connecteurs
     * {@code curvedConnector3} - confirmee correcte sur un rendu reel. */
    private static final double FRACTION_FROM_TIP_MULTI_SEGMENT = 0.40d;

    /** [v5-fraction-adaptative] Valeurs de fraction journalisees a titre diagnostique
     * pour chaque pointe recalculee - voir la Javadoc de
     * {@link #FRACTION_FROM_TIP_SINGLE_SEGMENT}. */
    private static final double[] DEBUG_FRACTION_SWEEP =
            {0.02d, 0.05d, 0.10d, 0.15d, 0.20d, 0.30d, 0.50d, 1.00d};

    /**
     * [v16-stealth-polygone] Position, en fraction de la longueur totale de la
     * pointe (0 = la pointe elle-meme, 1 = le bord arriere/barbes), du sommet
     * concave (l'encoche) d'une decoration {@code STEALTH} - voir
     * {@link TangentCorrectedConnector#buildDecoration}, cas {@code STEALTH}.
     * Valeur par defaut choisie visuellement (0.5, encoche a mi-longueur) pour
     * obtenir une silhouette "flechette" plausible et comparable a celle de
     * PowerPoint - non mesuree au pixel pres (contrairement aux autres
     * constantes de cette classe) : aucune reference PowerPoint fiable au
     * pixel pres n'etait disponible pour cette forme precise (contrairement a
     * l'angle de pointe des connecteurs courbes/coudes, mesure sur un rendu
     * reel - voir {@link #FRACTION_FROM_TIP_SINGLE_SEGMENT}). A ajuster si le
     * rendu reel s'avere visiblement trop/pas assez concave.
     */
    private static final double STEALTH_NOTCH_FRACTION = 0.5d;

    private ConnectorArrowFixer() {
    }

    /**
     * [v13-drawfactory-zorder] A appeler juste avant {@code slide.draw(graphics)}.
     * Installe un {@link DrawFactory} personnalise sur {@code graphics}, via le
     * hint de rendu {@code Drawable.DRAW_FACTORY}, afin que les connecteurs
     * courbes/coudes a pointe de fleche declaree soient dessines par
     * {@link TangentCorrectedConnector} au lieu du dessinateur standard de POI -
     * voir la Javadoc de la classe pour le detail du mecanisme et pourquoi il
     * preserve nativement l'ordre d'empilement (z-order) des formes.
     *
     * @return la valeur precedente du hint {@code Drawable.DRAW_FACTORY} sur ce
     * {@code Graphics2D} (peut etre {@code null}), a repasser telle quelle a
     * {@link #restoreAfterDraw} une fois {@code slide.draw(graphics)} termine.
     */
    public static Object installBeforeDraw(Graphics2D graphics) {
        Object previous = graphics.getRenderingHint(Drawable.DRAW_FACTORY);
        graphics.setRenderingHint(Drawable.DRAW_FACTORY, new ArrowCorrectingDrawFactory());
        if (LOG.isDebugEnabled()) {
            LOG.debug("[v13-drawfactory-zorder] DrawFactory de correction des pointes de fleche installe");
        }
        return previous;
    }

    /**
     * A appeler juste apres {@code slide.draw(graphics)}, avec la valeur
     * renvoyee par {@link #installBeforeDraw} - remet {@code graphics} dans
     * l'etat ou {@link #installBeforeDraw} l'a trouve (bonne pratique : evite
     * de laisser fuiter notre {@code DrawFactory} vers un dessin ulterieur sur
     * ce meme {@code Graphics2D} qui n'attendrait pas ce comportement).
     */
    public static void restoreAfterDraw(Graphics2D graphics, Object previousDrawFactory) {
        graphics.setRenderingHint(Drawable.DRAW_FACTORY, previousDrawFactory);
    }

    /**
     * {@code true} si {@code connector} est un connecteur courbe/coude (jamais
     * {@code straightConnector1}) qui declare une pointe de fleche (tete et/ou
     * queue) exploitable - meme condition que celle utilisee par les versions
     * v9-v12 pour decider de la correction.
     */
    private static boolean qualifiesForCorrection(XSLFConnectorShape connector) {
        if (connector.getShapeType() == ShapeType.STRAIGHT_CONNECTOR_1) {
            // Deja correct (voir Javadoc de la classe) : la ligne EST la diagonale
            // de la boite englobante, exactement ce que suppose le calcul de POI.
            return false;
        }
        LineDecoration deco = connector.getLineDecoration();
        if (deco == null) {
            return false;
        }
        DecorationShape headShape = deco.getHeadShape();
        DecorationShape tailShape = deco.getTailShape();
        boolean hasHead = headShape != null && headShape != DecorationShape.NONE;
        boolean hasTail = tailShape != null && tailShape != DecorationShape.NONE;
        if (!hasHead && !hasTail) {
            return false; // rien a corriger : pas de pointe de fleche declaree
        }
        Object xmlObject = connector.getXmlObject();
        if (!(xmlObject instanceof CTConnector)) {
            return false;
        }
        CTConnector ctConnector = (CTConnector) xmlObject;
        // spPr est un element obligatoire de cxnSp (comme bodyPr pour txBody, voir
        // RoundedShapeAnchorFixer) : XMLBeans ne genere donc pas de isSetSpPr() pour
        // lui, seulement pour ses propres elements optionnels comme "ln" ci-dessous.
        CTShapeProperties spPr = ctConnector.getSpPr();
        return spPr != null && spPr.isSetLn();
    }

    /**
     * [v13-drawfactory-zorder] {@code DrawFactory} personnalise qui substitue
     * {@link TangentCorrectedConnector} pour les connecteurs qualifies (voir
     * {@link #qualifiesForCorrection}), et delegue au comportement standard de
     * POI ({@code super.getDrawable(...)}) pour tout le reste - voir Javadoc
     * de la classe englobante.
     */
    private static final class ArrowCorrectingDrawFactory extends DrawFactory {
        @Override
        public DrawConnectorShape getDrawable(ConnectorShape<?, ?> shape) {
            if (shape instanceof XSLFConnectorShape) {
                XSLFConnectorShape connector = (XSLFConnectorShape) shape;
                if (qualifiesForCorrection(connector)) {
                    LineDecoration deco = connector.getLineDecoration();
                    return new TangentCorrectedConnector(connector,
                            deco.getHeadShape(), deco.getHeadWidth(), deco.getHeadLength(),
                            deco.getTailShape(), deco.getTailWidth(), deco.getTailLength());
                }
            }
            return super.getDrawable(shape);
        }
    }

    /**
     * Dessinateur de remplacement pour un connecteur courbe/coude a pointe
     * declaree - sous-classe de {@code DrawConnectorShape} (donc de
     * {@code DrawSimpleShape}), seul moyen d'acceder a
     * {@code computeOutlines()}/{@code getStroke()}/{@code getLinePaint()}
     * ({@code protected}) pour lire le trace et le style reels du connecteur.
     * Installe par {@link ArrowCorrectingDrawFactory}, appele par POI
     * exactement comme n'importe quel autre dessinateur (voir Javadoc de la
     * classe englobante) - {@link #draw} est donc le seul point d'entree,
     * invoque par POI APRES qu'il a lui-meme applique la transformation propre
     * a ce connecteur (rotation/{@code flipH}/{@code flipV}) sur
     * {@code graphics} : tout le calcul ci-dessous opere donc directement dans
     * ce repere deja en place, sans jamais avoir besoin d'appeler
     * {@code applyTransform} nous-memes, ni de projeter quoi que ce soit vers
     * un espace "device" pour un usage differe (contrairement aux versions
     * v9-v12).
     */
    private static final class TangentCorrectedConnector extends DrawConnectorShape {
        private final XSLFConnectorShape connector;
        private final DecorationShape headShape;
        private final DecorationSize headWidth;
        private final DecorationSize headLength;
        private final DecorationShape tailShape;
        private final DecorationSize tailWidth;
        private final DecorationSize tailLength;

        TangentCorrectedConnector(XSLFConnectorShape connector,
                                   DecorationShape headShape, DecorationSize headWidth, DecorationSize headLength,
                                   DecorationShape tailShape, DecorationSize tailWidth, DecorationSize tailLength) {
            super(connector);
            this.connector = connector;
            this.headShape = headShape;
            this.headWidth = headWidth;
            this.headLength = headLength;
            this.tailShape = tailShape;
            this.tailWidth = tailWidth;
            this.tailLength = tailLength;
        }

        /**
         * Reproduit la sequence de {@code DrawSimpleShape.draw()} (source
         * verifiee - voir Javadoc de la classe englobante), a l'identique pour
         * l'ombre et le trait, mais substitue {@link #buildDecoration} au
         * calcul de pointe standard de POI (bugue - voir Javadoc de la
         * classe englobante), et exclut localement, uniquement autour de
         * l'appel a {@code graphics.draw(...)} du trait, la zone pres de la
         * pointe pour une decoration {@code TRIANGLE} (voir
         * {@link #computeShaftExclusion}).
         */
        @Override
        public void draw(Graphics2D graphics) {
            Paint oldPaint = graphics.getPaint();
            Stroke oldStroke = graphics.getStroke();
            Color oldColor = graphics.getColor();
            Shape oldClip = graphics.getClip();

            Paint line = getLinePaint(graphics);
            BasicStroke stroke = getStroke();
            Collection<Outline> elems = computeOutlines(graphics);

            // [v14-detection-remplissage-corrigee] getFillPaint(graphics) n'est PAS un bon
            // signal de "ce connecteur a un remplissage visible" : pour un connecteur tout a
            // fait normal, sans aucun a:solidFill/a:noFill explicite dans son XML, il renvoie
            // presque toujours un Paint non-null (valeur heritee/par defaut du theme) - le
            // verifier faisait donc retomber la quasi-totalite des connecteurs sur le rendu
            // standard de POI (bug d'orientation inclus), constate sur un rendu reel apres
            // livraison de la v13 (tous les connecteurs a corriger y sont passes par
            // "remplissage detecte (cas non gere)"). Le signal fiable, deja utilise
            // par DrawSimpleShape.draw() lui-meme pour decider s'il y a quelque chose a
            // remplir, est de verifier si au moins un des Outline calcules par
            // computeOutlines() a un Path marque REELLEMENT rempli (Path.isFilled()) - un
            // simple trait de connecteur n'en a jamais (Path.isStroked() seul) : voir
            // Javadoc de la classe englobante, "Cas volontairement non gere".
            boolean hasFilledOutline = false;
            for (Outline o : elems) {
                if (o.getPath().isFilled()) {
                    hasFilledOutline = true;
                    break;
                }
            }

            if (line == null || stroke == null || elems.isEmpty() || hasFilledOutline) {
                // Cas 1: trace introuvable/degenere, ou style de trait non exploitable - on
                // renonce plutot que de deviner. Cas 2 (hasFilledOutline) : jamais rencontre en
                // pratique - voir Javadoc de la classe englobante. Dans les deux cas : rendu
                // standard de POI (potentiellement encore mal oriente, mais jamais pire que
                // sans ce correctif).
                if (LOG.isDebugEnabled() && hasFilledOutline) {
                    LOG.debug("[v14-detection-remplissage-corrigee] {} : remplissage reellement "
                                    + "detecte sur le trace lui-meme (cas non gere), rendu standard "
                                    + "POI utilise pour ce connecteur",
                            connector.getShapeName());
                }
                super.draw(graphics);
                return;
            }
            graphics.setStroke(stroke);

            try {
                drawShadow(graphics, elems, null, line);

                drawContent(graphics); // no-op herite (pas de contenu texte sur un connecteur)

                graphics.setPaint(line);
                graphics.setStroke(stroke);

                Shape headExclusion = computeShaftExclusion(graphics, true, stroke);
                Shape tailExclusion = computeShaftExclusion(graphics, false, stroke);
                if (headExclusion != null || tailExclusion != null) {
                    Area clip = oldClip != null
                            ? new Area(oldClip)
                            : new Area(new Rectangle2D.Double(-1_000_000, -1_000_000, 2_000_000, 2_000_000));
                    if (headExclusion != null) {
                        clip.subtract(new Area(headExclusion));
                    }
                    if (tailExclusion != null) {
                        clip.subtract(new Area(tailExclusion));
                    }
                    graphics.setClip(clip);
                }

                for (Outline o : elems) {
                    if (o.getPath().isStroked()) {
                        Shape s = o.getOutline();
                        graphics.setRenderingHint(Drawable.GRADIENT_SHAPE, s);
                        graphics.draw(s);
                    }
                }

                graphics.setClip(oldClip);

                drawCorrectedDecoration(graphics, line, stroke);
            } finally {
                graphics.setClip(oldClip);
                graphics.setColor(oldColor);
                graphics.setPaint(oldPaint);
                graphics.setStroke(oldStroke);
            }
        }

        /**
         * Remplace {@code DrawSimpleShape.drawDecoration} : meme boucle
         * fill/stroke par {@code Outline}, mais a partir de
         * {@link #buildDecoration} (angle/position corriges) plutot que de
         * {@code getHeadDecoration}/{@code getTailDecoration} (bugues).
         */
        private void drawCorrectedDecoration(Graphics2D graphics, Paint line, BasicStroke stroke) {
            if (line == null) {
                return;
            }
            graphics.setPaint(line);

            Outline head = buildDecoration(graphics, stroke, true);
            Outline tail = buildDecoration(graphics, stroke, false);
            for (Outline o : new Outline[]{head, tail}) {
                if (o == null) {
                    continue;
                }
                Shape s = o.getOutline();
                graphics.setRenderingHint(Drawable.GRADIENT_SHAPE, s);
                if (o.getPath().isFilled()) {
                    graphics.fill(s);
                }
                if (o.getPath().isStroked()) {
                    graphics.setStroke(stroke);
                    graphics.draw(s);
                }
            }
        }

        /**
         * [v13-drawfactory-zorder] Calcule, dans le repere DEJA en place au
         * moment de l'appel (voir Javadoc de la classe : celui applique par
         * POI juste avant d'invoquer {@link #draw}, donc le meme que celui
         * utilise par {@code computeOutlines()}/{@code graphics.draw(...)}
         * plus haut), la zone d'exclusion pour une seule extremite
         * ({@code head}/{@code tail}) : un petit rectangle, dans le repere
         * local du connecteur (apex a l'origine, meme convention +x/-x que
         * {@link #buildDecoration}), couvrant la moitie de la longueur de la
         * pointe la plus proche de son extremite - meme longueur que celle
         * mesuree sur un rendu reel (voir Javadoc de la classe englobante,
         * "trait raccourci pres de la pointe"). Contrairement a v11/v12, ce
         * rectangle n'est PAS projete vers un espace "device" separe : il est
         * directement dans le repere courant, pret a etre utilise comme clip.
         *
         * @return {@code null} si la decoration de cette extremite n'est pas
         * {@code TRIANGLE} (seule forme presentant l'inadequation de largeur -
         * voir Javadoc de la classe englobante), ou si aucun trace exploitable
         * n'a ete trouve pour cette pointe.
         */
        private Shape computeShaftExclusion(Graphics2D graphics, boolean head, BasicStroke stroke) {
            DecorationShape decoShape = head ? headShape : tailShape;
            // [v16-stealth-polygone] Meme inadequation de largeur que pour TRIANGLE
            // (voir Javadoc de la classe englobante, "trait raccourci pres de la
            // pointe") : STEALTH est desormais aussi un polygone plein qui s'amincit
            // jusqu'a une largeur nulle a la pointe (voir buildDecoration) - le trait
            // brut, de largeur constante, deborderait donc sur ses cotes pres de la
            // pointe exactement de la meme facon que pour TRIANGLE sans cette exclusion.
            if (decoShape != DecorationShape.TRIANGLE && decoShape != DecorationShape.STEALTH) {
                return null;
            }
            DecorationSize length = head ? headLength : tailLength;
            if (length == null) {
                length = DecorationSize.MEDIUM;
            }

            double[] tip = computeEndpoint(graphics, head);
            if (tip == null) {
                return null;
            }
            double tipX = tip[0], tipY = tip[1], alpha = tip[2];
            double lineWidth = Math.max(2.5, stroke.getLineWidth());
            double scaleX = Math.pow(DECO_SIZE_POW, length.ordinal() + 1.);
            double arrowLength = lineWidth * scaleX;
            double shaftWidth = stroke.getLineWidth();

            double excludeLength = arrowLength / 2.0;
            double sign = head ? 1 : -1;
            double rectX = sign < 0 ? -excludeLength : 0;
            // Marge au-dela de la largeur du trait, pour absorber sans risque un
            // leger deborder (ex. extremite CAP_SQUARE).
            double rectHeight = shaftWidth + 2.0;

            AffineTransform local = new AffineTransform();
            local.translate(tipX, tipY);
            local.rotate(alpha);
            Rectangle2D.Double localRect =
                    new Rectangle2D.Double(rectX, -rectHeight / 2.0, excludeLength, rectHeight);
            Shape exclusion = local.createTransformedShape(localRect);

            if (LOG.isDebugEnabled()) {
                LOG.debug("[v13-drawfactory-zorder] {} : {} zone d'exclusion - longueur {} pt "
                                + "(moitie pointe, longueur totale {} pt), hauteur {} pt, pointe a "
                                + "({}, {}) pt angle {} deg",
                        connector.getShapeName(), head ? "tete" : "queue",
                        String.format(Locale.ROOT, "%.2f", excludeLength),
                        String.format(Locale.ROOT, "%.2f", arrowLength),
                        String.format(Locale.ROOT, "%.2f", rectHeight),
                        String.format(Locale.ROOT, "%.2f", tipX), String.format(Locale.ROOT, "%.2f", tipY),
                        String.format(Locale.ROOT, "%.1f", Math.toDegrees(alpha)));
            }

            return exclusion;
        }

        /**
         * Construit la geometrie locale de la pointe {@code head}/{@code tail},
         * avec un angle/position corriges. Reproduit la meme construction
         * geometrique de pointe de fleche (ovale/fleche simple/triangle) que
         * {@code DrawSimpleShape} - la marche a suivre (translation puis
         * rotation d'un petit gabarit local) est fidele au code de POI, seul le
         * calcul de l'angle et du point d'ancrage change (voir Javadoc de la
         * classe englobante).
         */
        private Outline buildDecoration(Graphics2D graphics, BasicStroke stroke, boolean head) {
            DecorationShape decoShape = head ? headShape : tailShape;
            if (decoShape == null || decoShape == DecorationShape.NONE) {
                return null;
            }
            DecorationSize width = head ? headWidth : tailWidth;
            DecorationSize length = head ? headLength : tailLength;
            if (width == null) {
                width = DecorationSize.MEDIUM;
            }
            if (length == null) {
                length = DecorationSize.MEDIUM;
            }

            double[] tip = computeEndpoint(graphics, head);
            if (tip == null) {
                return null; // trace introuvable/degenere : on renonce plutot que de deviner
            }
            double tipX = tip[0], tipY = tip[1], alpha = tip[2];
            double lineWidth = Math.max(2.5, stroke.getLineWidth());

            if (LOG.isDebugEnabled()) {
                Rectangle2D declaredAnchor = connector.getAnchor();
                LOG.debug("{} : {} recalculee - decoShape={} width={} "
                                + "length={} lineWidth={} pt - sous-corde a {} de la pointe : "
                                + "({};{}) -> ({};{}) pt, pointe finale a "
                                + "({}, {}) pt, angle {} deg (ancre declaree : {})",
                        connector.getShapeName(), head ? "tete" : "queue",
                        decoShape, width, length, String.format(Locale.ROOT, "%.2f", lineWidth),
                        String.format(Locale.ROOT, "%.0f%%", tip[7] * 100),
                        String.format(Locale.ROOT, "%.2f", tip[3]), String.format(Locale.ROOT, "%.2f", tip[4]),
                        String.format(Locale.ROOT, "%.2f", tip[5]), String.format(Locale.ROOT, "%.2f", tip[6]),
                        String.format(Locale.ROOT, "%.2f", tipX), String.format(Locale.ROOT, "%.2f", tipY),
                        String.format(Locale.ROOT, "%.1f", Math.toDegrees(alpha)), declaredAnchor);
            }
            double scaleY = Math.pow(DECO_SIZE_POW, width.ordinal() + 1.);
            double scaleX = Math.pow(DECO_SIZE_POW, length.ordinal() + 1.);
            // Gabarit local : la tete pointe vers +x (base en +x), la queue vers -x
            // (base en -x) - meme convention que DrawSimpleShape.
            double sign = head ? 1 : -1;

            AffineTransform at = new AffineTransform();
            Shape decoGeom;
            Path p;

            switch (decoShape) {
                case OVAL: {
                    p = new Path();
                    // [v15-sommets-arrondis] Forme pleine : contour NON trace separement,
                    // voir la note complete au cas TRIANGLE ci-dessous (meme raison, ici
                    // l'effet est juste moins visible car un cercle n'a pas d'angle vif).
                    p.setStroke(false);
                    Ellipse2D oval = new Ellipse2D.Double(0, 0, lineWidth * scaleX, lineWidth * scaleY);
                    Rectangle2D bounds = oval.getBounds2D();
                    at.translate(tipX - bounds.getWidth() / 2, tipY - bounds.getHeight() / 2);
                    at.rotate(alpha, bounds.getX() + bounds.getWidth() / 2, bounds.getY() + bounds.getHeight() / 2);
                    decoGeom = oval;
                    break;
                }
                case ARROW: {
                    // "Fleche ouverte" : un simple chevron trace au trait (pas rempli),
                    // fidele a la fois au rendu de POI et a celui de PowerPoint pour ce
                    // type precis - contrairement a STEALTH ci-dessous, PowerPoint la
                    // dessine bien comme un trait, pas comme une forme pleine, donc pas
                    // du meme defaut "sommets emousses" (voir Javadoc [v16-stealth-polygone]).
                    p = new Path();
                    p.setFill(PaintModifier.NONE);
                    p.setStroke(true);
                    Path2D.Double arrow = new Path2D.Double();
                    arrow.moveTo(sign * lineWidth * scaleX, -lineWidth * scaleY / 2);
                    arrow.lineTo(0, 0);
                    arrow.lineTo(sign * lineWidth * scaleX, lineWidth * scaleY / 2);
                    decoGeom = arrow;
                    at.translate(tipX, tipY);
                    at.rotate(alpha);
                    break;
                }
                case STEALTH: {
                    // [v16-stealth-polygone] Corrige un ecart de fidelite face au rendu
                    // PowerPoint natif (pointe "flechette" a barbes/encoche concave) :
                    // POI (et notre code jusqu'a present) dessinait cette pointe
                    // exactement comme ARROW ci-dessus - un simple chevron trace au
                    // trait, pas une forme pleine. Un trait, meme fin, ne peut pas
                    // produire les sommets nets d'un vrai polygone (le sommet central
                    // s'arrondit/s'aplatit selon la jointure du trait, et les deux
                    // extremites OUVERTES du chevron ne sont que des bouts de trait -
                    // arrondis ou plats selon le "cap" - jamais un vrai angle rentrant ni
                    // les barbes pointues que montre PowerPoint) : confirme par comparaison
                    // directe avec une capture PowerPoint (sommets nets des 2 cotes) contre
                    // le rendu produit ici (sommets systematiquement emousses, y compris
                    // l'encoche concave qui apparaissait juste comme une bosse arrondie).
                    // Corrige en construisant un vrai polygone plein a 4 sommets (pointe,
                    // barbe, encoche concave, barbe) - voir STEALTH_NOTCH_FRACTION pour la
                    // position de l'encoche. setStroke(false) est indispensable ici : voir
                    // la lecon du [v15-sommets-arrondis] plus haut - un Path() par defaut a
                    // isStroked()==true, et un second trace par-dessus le remplissage
                    // emousserait de nouveau tous les sommets de ce polygone, exactement
                    // comme c'etait le cas pour TRIANGLE avant ce correctif.
                    p = new Path();
                    p.setStroke(false);
                    double stealthLength = lineWidth * scaleX;
                    double stealthWidth = lineWidth * scaleY;
                    Path2D.Double stealth = new Path2D.Double();
                    stealth.moveTo(0, 0);
                    stealth.lineTo(sign * stealthLength, -stealthWidth / 2);
                    stealth.lineTo(sign * stealthLength * STEALTH_NOTCH_FRACTION, 0);
                    stealth.lineTo(sign * stealthLength, stealthWidth / 2);
                    stealth.closePath();
                    decoGeom = stealth;
                    at.translate(tipX, tipY);
                    at.rotate(alpha);
                    break;
                }
                case TRIANGLE: {
                    p = new Path();
                    // [v15-sommets-arrondis] Path() par defaut a stroke=true (verifie dans
                    // le source POI 5.2.5, org.apache.poi.sl.draw.geom.Path - isStroked()
                    // vaut true tant que setStroke(false) n'est pas appele explicitement),
                    // ce qui restait invisible ici jusqu'a present car personne ne l'avait
                    // desactive pour OVAL/TRIANGLE (seuls STEALTH/ARROW le configurent
                    // explicitement, a true, car ce sont des chevrons non remplis). Consequence
                    // concrete dans drawCorrectedDecoration() : le triangle etait REMPLI
                    // (isFilled()==true, correct) PUIS son CONTOUR retrace par-dessus avec le
                    // stroke du connecteur (jointures/extremites arrondies) - un simple
                    // rehaussement de bordure pour une grande pointe, mais qui "mange"
                    // visiblement les 3 sommets d'une petite pointe des lors que la largeur du
                    // trait n'est plus negligeable devant la taille du triangle (l'effet est
                    // proportionnel a la pointe, pas a une valeur absolue - confirme par
                    // reproduction isolee, java.awt seul, hors POI : a largeur de trait fixe,
                    // une pointe de 12px ressort nettement arrondie sur ses 3 sommets, une
                    // pointe de 100px reste anguleuse). C'est le defaut releve par comparaison
                    // avec le rendu PowerPoint (sommets arrondis vs pointes anguleuses attendues).
                    // Une pointe de fleche PLEINE n'a pas besoin d'un contour trace en plus de
                    // son remplissage (contrairement a un chevron ARROW/STEALTH, non rempli,
                    // qui EST son propre contour) : setStroke(false) supprime ce second passage
                    // et restaure des sommets nets a toute taille.
                    p.setStroke(false);
                    Path2D.Double triangle = new Path2D.Double();
                    triangle.moveTo(sign * lineWidth * scaleX, -lineWidth * scaleY / 2);
                    triangle.lineTo(0, 0);
                    triangle.lineTo(sign * lineWidth * scaleX, lineWidth * scaleY / 2);
                    triangle.closePath();
                    // Le trait "brut" ne deborde plus sous ce triangle : voir
                    // TangentCorrectedConnector#draw, qui exclut deja localement la zone
                    // concernee lors du dessin du trait (voir Javadoc de la classe
                    // englobante, section "trait raccourci pres de la pointe") - pas
                    // besoin de recouvrement ici.
                    decoGeom = triangle;

                    at.translate(tipX, tipY);
                    at.rotate(alpha);
                    break;
                }
                default:
                    // DIAMOND : POI lui-meme ne le dessine pas non plus (meme "default: break;"
                    // dans le code d'origine) - non traite ici pour la meme raison.
                    return null;
            }

            Shape transformed = at.createTransformedShape(decoGeom);
            return new Outline(transformed, p);
        }

        /**
         * Renvoie {@code {x, y, alpha, subX, subY, tipX, tipY, fraction}} : le
         * point d'extremite reel du trace (pas le coin de la boite englobante,
         * contrairement au code d'origine de POI - verifie correct a &lt;4px
         * pres par mesure pixel) et l'angle d'une
         * <b>sous-corde proche de la pointe</b> du premier/dernier segment de
         * trace : le segment terminal est parametre en {@code t} de 0 (son
         * origine) a 1 (son extremite), et on prend la corde entre le point a
         * une fraction donnee de la pointe et l'extremite elle-meme - la
         * fraction utilisee depend du nombre total de segments du trace, voir
         * {@link #FRACTION_FROM_TIP_SINGLE_SEGMENT}/
         * {@link #FRACTION_FROM_TIP_MULTI_SEGMENT} pour la justification de ce
         * choix plutot qu'une fraction fixe unique, la corde du segment entier
         * ou la tangente instantanee au point de controle.
         *
         * <p>Le point intermediaire est obtenu par l'algorithme de De Casteljau,
         * applicable uniformement aux segments {@code lineTo} (degre 1),
         * {@code quadTo} (degre 2) et {@code cubicTo} (degre 3) rencontres dans
         * les traces de connecteurs.
         *
         * @return {@code null} si le trace est introuvable ou degenere (segment
         * de longueur nulle) - dans ce cas, aucune pointe n'est dessinee plutot
         * que de deviner une orientation.
         */
        private double[] computeEndpoint(Graphics2D graphics, boolean head) {
            Collection<Outline> outlines = computeOutlines(graphics);
            for (Outline outline : outlines) {
                double[] coords = new double[6];
                double curX = Double.NaN, curY = Double.NaN; // point courant pendant le parcours

                boolean haveFirstSeg = false;
                double[][] firstSegPts = null; // control points, P0..Pn, du premier segment

                boolean haveLastSeg = false;
                double[][] lastSegPts = null; // control points, P0..Pn, du dernier segment

                int segIndex = 0;
                PathIterator it = outline.getOutline().getPathIterator(null);
                while (!it.isDone()) {
                    int type = it.currentSegment(coords);
                    double[][] segPts;
                    switch (type) {
                        case PathIterator.SEG_MOVETO:
                            curX = coords[0];
                            curY = coords[1];
                            it.next();
                            continue;
                        case PathIterator.SEG_LINETO:
                            segPts = new double[][]{{curX, curY}, {coords[0], coords[1]}};
                            break;
                        case PathIterator.SEG_QUADTO:
                            segPts = new double[][]{{curX, curY}, {coords[0], coords[1]}, {coords[2], coords[3]}};
                            break;
                        case PathIterator.SEG_CUBICTO:
                            segPts = new double[][]{{curX, curY}, {coords[0], coords[1]},
                                    {coords[2], coords[3]}, {coords[4], coords[5]}};
                            break;
                        default:
                            // SEG_CLOSE : un trace de connecteur est toujours ouvert en
                            // pratique - ignore par prudence plutot que suppose.
                            it.next();
                            continue;
                    }
                    segIndex++;
                    double endX = segPts[segPts.length - 1][0];
                    double endY = segPts[segPts.length - 1][1];
                    if (!haveFirstSeg) {
                        firstSegPts = segPts;
                        haveFirstSeg = true;
                    }
                    lastSegPts = segPts;
                    haveLastSeg = true;
                    curX = endX;
                    curY = endY;
                    it.next();
                }

                // [v5-fraction-adaptative] la fraction a utiliser depend du nombre total
                // de segments du trace (voir Javadoc de FRACTION_FROM_TIP_SINGLE_SEGMENT) -
                // segIndex, a la fin de la boucle ci-dessus, vaut ce total.
                double fraction = segIndex <= 1 ? FRACTION_FROM_TIP_SINGLE_SEGMENT : FRACTION_FROM_TIP_MULTI_SEGMENT;
                if (head && haveFirstSeg) {
                    double[] result = tangentNearTip(firstSegPts, true, "tete", fraction, segIndex);
                    if (result != null) {
                        return result;
                    }
                }
                if (!head && haveLastSeg) {
                    double[] result = tangentNearTip(lastSegPts, false, "queue", fraction, segIndex);
                    if (result != null) {
                        return result;
                    }
                }
            }
            return null;
        }

        /**
         * Calcule la sous-corde proche de la pointe pour un segment donne (voir
         * Javadoc de {@link #computeEndpoint}), et journalise au passage le
         * balayage {@link #DEBUG_FRACTION_SWEEP} pour permettre d'ajuster
         * {@link #FRACTION_FROM_TIP_SINGLE_SEGMENT}/
         * {@link #FRACTION_FROM_TIP_MULTI_SEGMENT} sans nouvelle mesure a
         * l'aveugle.
         *
         * @param segPts points de controle du segment, {@code P0..Pn} dans
         *               l'ordre du trace (P0 = origine du segment, Pn = son
         *               extremite)
         * @param forward {@code true} pour une tete de fleche (la pointe est a
         *                l'origine P0 du segment, l'axe se mesure en s'eloignant
         *                de P0 vers l'interieur du segment) ; {@code false} pour
         *                une queue de fleche (la pointe est a l'extremite Pn).
         * @param fraction fraction (0-1) du segment, mesuree depuis la pointe,
         *                 utilisee pour la sous-corde - voir
         *                 {@link #FRACTION_FROM_TIP_SINGLE_SEGMENT}.
         * @param totalSegmentCount nombre total de segments du trace (a but
         *                          diagnostique uniquement, voir le message
         *                          journalise ci-dessous).
         * @return {@code {x, y, alpha, subX, subY, tipX, tipY, fraction}} ou
         * {@code null} si le segment est degenere (sous-corde de longueur nulle)
         */
        private double[] tangentNearTip(double[][] segPts, boolean forward, String label, double fraction,
                                         int totalSegmentCount) {
            double[] tipPt = forward ? segPts[0] : segPts[segPts.length - 1];
            if (LOG.isDebugEnabled()) {
                StringBuilder sweep = new StringBuilder();
                for (double f : DEBUG_FRACTION_SWEEP) {
                    double t = forward ? f : 1d - f;
                    double[] sample = deCasteljau(segPts, t);
                    double sx = forward ? tipPt[0] : sample[0];
                    double sy = forward ? tipPt[1] : sample[1];
                    double ex = forward ? sample[0] : tipPt[0];
                    double ey = forward ? sample[1] : tipPt[1];
                    double angDeg = Math.toDegrees(Math.atan2(ey - sy, ex - sx));
                    sweep.append(String.format(Locale.ROOT, " t=%.0f%%->%.1fdeg", f * 100, angDeg));
                }
                LOG.debug("  {} ({} segment(s) dans le trace, fraction "
                                + "retenue {}) balayage angle(sous-corde) en fonction de la fraction "
                                + "depuis la pointe :{}",
                        label, totalSegmentCount, String.format(Locale.ROOT, "%.0f%%", fraction * 100), sweep);
            }
            double t = forward ? fraction : 1d - fraction;
            double[] sub = deCasteljau(segPts, t);
            double startX = forward ? tipPt[0] : sub[0];
            double startY = forward ? tipPt[1] : sub[1];
            double endX = forward ? sub[0] : tipPt[0];
            double endY = forward ? sub[1] : tipPt[1];
            double tx = endX - startX;
            double ty = endY - startY;
            if (tx == 0 && ty == 0) {
                return null; // sous-corde degeneree : on renonce plutot que de deviner
            }
            return new double[]{tipPt[0], tipPt[1], Math.atan2(ty, tx), startX, startY, endX, endY, fraction};
        }

        /**
         * Algorithme de De Casteljau, applicable uniformement a un segment de
         * degre 1 ({@code lineTo}, 2 points), 2 ({@code quadTo}, 3 points) ou 3
         * ({@code cubicTo}, 4 points) : renvoie le point du segment au
         * parametre {@code t} (0 = origine du segment, 1 = son extremite).
         */
        private static double[] deCasteljau(double[][] pts, double t) {
            double[][] cur = pts;
            while (cur.length > 1) {
                double[][] next = new double[cur.length - 1][2];
                for (int i = 0; i < next.length; i++) {
                    next[i][0] = cur[i][0] + (cur[i + 1][0] - cur[i][0]) * t;
                    next[i][1] = cur[i][1] + (cur[i + 1][1] - cur[i][1]) * t;
                }
                cur = next;
            }
            return cur[0];
        }
    }
}
