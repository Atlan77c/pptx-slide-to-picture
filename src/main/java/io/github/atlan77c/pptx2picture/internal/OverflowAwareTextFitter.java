package io.github.atlan77c.pptx2picture.internal;

import org.apache.poi.sl.usermodel.ShapeType;
import org.apache.poi.sl.usermodel.TextShape;
import org.apache.poi.sl.usermodel.VerticalAlignment;
import org.apache.poi.xslf.model.PropertyFetcher;
import org.apache.poi.xslf.usermodel.XSLFGroupShape;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextBox;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextBody;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextBodyProperties;
import org.openxmlformats.schemas.presentationml.x2006.main.CTShape;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.awt.Font;
import java.awt.FontMetrics;
import java.awt.Graphics2D;
import java.awt.font.FontRenderContext;
import java.awt.font.TextLayout;
import java.awt.geom.Rectangle2D;
import java.text.Normalizer;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Corrige un ecart de fidelite observe entre le rendu natif d'Apache POI
 * ({@code XSLFSlide.draw(Graphics2D)}) et le rendu natif de PowerPoint :
 * pour certaines polices, les metriques verticales calculees par le pipeline
 * Java2D/AWT surestiment la hauteur de texte reellement necessaire (jusqu'a
 * 30-35% observe dans certains cas) par rapport a ce que produit le moteur de
 * rendu de PowerPoint. POI n'applique par ailleurs jamais le retrecissement
 * automatique de police ("shrink text on overflow", stocke par PowerPoint
 * dans {@code <a:normAutofit fontScale="...">}) lors du rendu :
 * {@code slide.draw()} utilise toujours la taille de police brute.
 *
 * <p>Consequence non corrigee : une zone de texte peut deborder visuellement
 * de sa boite d'origine et chevaucher une autre forme voisine.
 *
 * <p><b>Strategie retenue</b> : on ne "triche" pas sur les metriques de ligne calculees par
 * Java2D (une compensation directe de {@code DrawTextFragment.getHeight()}
 * corrige bien le debordement mais introduit un nouveau chevauchement entre
 * lignes consecutives d'une meme zone de texte, car cette valeur sert aussi
 * de pas d'avancement vertical entre lignes dans POI). On reduit a la place
 * la taille de police reelle des formes concernees, ce qui reduit glyphes et
 * interlignage ensemble et proportionnellement - par construction, cela ne
 * peut jamais faire chevaucher deux lignes entre elles.
 *
 * <p>Pour eviter de retrecir des formes qui n'en ont pas besoin (la
 * surestimation touche la quasi-totalite des formes d'un fichier utilisant
 * une police concernee, y compris des formes isolees sans aucun risque de
 * collision visuelle reelle), le retrecissement d'une forme marquee
 * {@link TextShape.TextAutofit#NONE} n'est declenche que si son debordement
 * calcule chevaucherait reellement l'emplacement d'une autre forme de texte
 * non vide du meme slide. Les formes {@link TextShape.TextAutofit#NORMAL}
 * (retrecissement automatique voulu par l'auteur du fichier) sont, elles,
 * toujours retrecies en cas de depassement reel - c'est le comportement
 * natif de PowerPoint pour ce type de forme.
 *
 * <p><b>Cas {@link TextShape.TextAutofit#SHAPE}</b> ("redimensionner la forme
 * selon le texte", {@code spAutoFit} en OOXML) : traite comme {@code NORMAL}
 * (retrecissement systematique de la police en cas de depassement mesure).
 * Une premiere version agrandissait plutot la hauteur de l'anchor - fidele au
 * sens strict de ce mode d'autofit dans PowerPoint (qui ne reduit jamais la
 * police, la boite grandit) - mais observee en pratique sur un fichier reel
 * a provoquer de nouveaux chevauchements avec les formes voisines situees en
 * dessous, puisque l'agrandissement deplace la limite basse de la boite sans
 * tenir compte du reste de la mise en page. Retrecir la police laisse toutes
 * les autres formes du slide a leur place d'origine : moins fidele au
 * mecanisme technique de ce mode d'autofit, mais plus fidele au rendu global
 * du diagramme, ce qui est le critere qui compte ici.
 *
 * <p><b>Exception au retrecissement force ({@code NONE})</b> : decouverte sur
 * un fichier reel (une boite {@code noAutofit} nettement plus petite que la
 * taille de police declaree, ne contenant qu'un seul caractere decoratif
 * volontairement surdimensionne par l'auteur pour dessiner un accent visuel -
 * chevauchant deliberement une forme voisine, exactement comme PowerPoint
 * l'affiche). Si la taille de police <em>declaree</em> (avant toute mesure)
 * depasse deja a elle seule la hauteur de la boite, le debordement n'est pas
 * un artefact de mesure Java2D : PowerPoint montrerait le meme debordement
 * quelle que soit la precision du calcul de metriques, puisque meme un calcul
 * parfait ne ferait jamais tenir un texte aussi grand dans une boite aussi
 * petite. Un tel cas n'est donc jamais retreci de force, meme s'il chevauche
 * effectivement une autre forme de texte.
 *
 * <p><b>Deuxieme garde-fou ({@code NONE}) : objectif "sans collision"
 * plutot que "tient dans la boite"</b> - decouverte sur un autre fichier
 * reel (boite {@code noAutofit} contenant plusieurs paragraphes de texte,
 * alignement "Milieu"). Contrairement au cas precedent, aucune taille de
 * police declaree ne depasse a elle seule la hauteur de la boite : c'est le
 * volume de texte (nombre de lignes) qui fait que le retrecissement visant
 * un ajustement complet dans la boite (le meme objectif que pour
 * {@code NORMAL}/{@code SHAPE}) n'atteint jamais sa cible, meme ecrase
 * jusqu'a la limite basse ({@link #MIN_SCALE}) ou {@link #MAX_ITER}
 * iterations. Or PowerPoint n'exige jamais qu'un texte {@code noAutofit}
 * tienne entierement dans sa boite - seulement qu'il ne produise pas de
 * chevauchement genant avec une forme voisine, ce qui est la seule raison
 * d'etre du retrecissement force ici. Quand l'ajustement complet echoue, une
 * deuxieme passe vise donc un objectif plus modeste : ne retrecir que
 * jusqu'a ce que la collision reelle detectee plus haut disparaisse
 * (recalculee a chaque iteration), en acceptant un debordement residuel
 * au-dela de la boite tant qu'il ne chevauche plus la forme voisine -
 * exactement ce que montre PowerPoint pour ce type de boite. Si meme cette
 * cible plus modeste n'est pas atteinte a la limite basse, la taille
 * d'origine est restauree en dernier recours plutot que de produire un
 * texte ecrase qui chevauche quand meme.
 *
 * <h2>Troisieme garde-fou ({@code NONE}) : paragraphe entierement vide
 * ignore dans la mesure (2026-08-29)</h2>
 * <p>Decouverte sur le slide 2 (sommaire) du fichier "Refonte BEL -
 * Trajectoire d'adressage - V.3 - Aout 2026" : un item du sommaire
 * ("Trajectoire d'adressage") porte, en plus de son unique paragraphe
 * visible, un SECOND paragraphe entierement vide ({@code endParaRPr} seul,
 * aucun {@code <a:r>}) - un Entree reste par megarde dans le fichier source,
 * sans equivalent visuel dans PowerPoint (rien ne se dessine pour un
 * paragraphe vide). {@code ts.getTextHeight(graphics)} mesure neanmoins ce
 * paragraphe fantome comme une ligne de texte reelle a part entiere, ce qui
 * double artificiellement la hauteur mesuree. Contrairement au cas d'un
 * texte trop long pour sa largeur (voir item 1 du meme slide) - ou reduire
 * la police reduit aussi la largeur et fait naturellement disparaitre le
 * repli a la ligne - un paragraphe vide ne "disparait" jamais en retrecissant
 * la police : {@link #captureBaselineFontSizes} ne capture que des
 * {@link XSLFTextRun} reels, et un paragraphe vide n'en contient aucun -
 * son propre retrecissement de police (via {@link #fitOverflowingText}) ne
 * touche donc jamais ce paragraphe, dont la hauteur (a la taille de police
 * heritee/par defaut du paragraphe) reste constante quel que soit le facteur
 * applique aux runs reels. Consequence observee : le texte visible etait
 * retreci bien au-dela de ce que sa propre collision reelle exigeait -
 * jusqu'a l'epuisement de {@link #MIN_SCALE} dans certains cas - puisque la
 * boucle de retrecissement (visant a faire tenir la hauteur MESUREE, paragraphe
 * fantome compris, dans la boite) ne pouvait structurellement jamais atteindre
 * sa cible tant que la contribution constante du paragraphe vide a elle seule
 * approchait ou depassait deja la hauteur de la boite.
 *
 * <p><b>Correction retenue</b> : pour une forme {@code noAutofit} melangeant
 * au moins un paragraphe visible et au moins un paragraphe entierement vide,
 * toute hauteur de texte utilisee pour la detection (zone de debordement
 * initiale) et le calibrage (boucles de retrecissement, objectif complet
 * comme objectif de repli "collision") est mesuree via {@link
 * VisibleTextMeasurer}, qui ignore les paragraphes entierement vides -
 * exactement comme {@code NeighborShapeOverlapFixer} ignore deja ces memes
 * paragraphes fantomes dans sa propre detection de collision par paragraphe.
 * Un paragraphe vide n'ayant aucun contenu visible a proteger, il n'a pas a
 * peser dans la decision de retrecir ou non le texte reel qui l'accompagne.
 * Portee volontairement limitee aux formes {@code noAutofit} : pour {@code
 * NORMAL}/{@code SHAPE} (retrecissement automatique voulu par l'auteur),
 * PowerPoint calcule reellement son {@code fontScale} sur la hauteur totale
 * du bloc, paragraphes vides compris - les en exclure ici introduirait un
 * nouvel ecart avec PowerPoint plutot que d'en corriger un.
 *
 * <p><b>Amendement (2026-09-01) : paragraphe vide EN MILIEU de bloc</b> - la
 * correction ci-dessus, validee uniquement sur le cas d'un paragraphe vide EN
 * FIN de bloc (aucun paragraphe visible apres lui), s'est averee incorrecte
 * pour un paragraphe vide EN MILIEU de bloc (au moins un paragraphe visible
 * le suit) : decouvert sur le slide 4 du fichier "fichier-test-A.pptx" (forme
 * "ZoneTexte 23", {@code noAutofit} explicitement declare, 9 paragraphes dont
 * deux entierement vides intercales parmi des paragraphes visibles). Un
 * paragraphe vide EN FIN de bloc ne "pousse" rien apres lui - rien ne depend
 * de sa hauteur reelle, l'ignorer entierement pour la mesure est donc sans
 * consequence sur le positionnement du texte reellement dessine. Un
 * paragraphe vide EN MILIEU de bloc, en revanche, occupe bel et bien un
 * espace reel et constant (non retrecissable, comme documente plus haut) qui
 * repousse vers le bas tous les paragraphes visibles qui le suivent dans le
 * rendu reel de POI - l'ignorer entierement pour la mesure sous-estime donc
 * la hauteur reellement necessaire, ce qui a entraine un retrecissement
 * insuffisant (96% seulement, mesure via debug log) et un nouveau
 * chevauchement reel avec une forme voisine ("Round Same Side Corner
 * Rectangle 1"/"ZoneTexte 21", l'en-tete "SOLUTION") qui n'existait pas avant
 * l'introduction de ce garde-fou le 2026-08-29 (confirme par une capture
 * d'ecran de reference anterieure a son introduction, fournie par
 * l'utilisateur).
 *
 * <p><b>Correction retenue</b> : {@link VisibleTextMeasurer} n'ignore plus
 * desormais qu'un paragraphe vide faisant partie de la SEQUENCE FINALE de
 * paragraphes vides du bloc (lui-meme vide, et tous les paragraphes qui le
 * suivent jusqu'a la fin du bloc egalement vides) - un paragraphe vide suivi,
 * plus loin dans le bloc, d'au moins un paragraphe visible reste desormais
 * inclus dans la mesure, avec sa hauteur reelle (a la taille de police
 * heritee/par defaut du paragraphe, comme documente plus haut), pour
 * preserver sa contribution reelle au positionnement du texte visible qui le
 * suit. Le declenchement de {@link VisibleTextMeasurer} lui-meme (voir {@code
 * ignoreBlankParagraphs} dans {@link #fitOverflowingText}) suit desormais la
 * meme logique : {@link #hasTrailingBlankParagraph} ne retourne {@code true}
 * que si le DERNIER paragraphe du bloc est entierement vide (ce qui implique
 * l'existence d'au moins une sequence finale de paragraphes vides a ignorer).
 * Pour "ZoneTexte 23" ci-dessus, dont le dernier paragraphe (le 9e) est
 * visible, cette condition est desormais fausse : la forme n'entre plus du
 * tout dans ce garde-fou, {@code ts.getTextHeight(graphics)} est utilisee
 * directement, exactement comme avant l'introduction de ce garde-fou le
 * 2026-08-29.
 *
 * <h2>Elargissement cible aux diapositives "sommaire"/"table des matieres"
 * (2026-08-31)</h2>
 * <p>Analyse plus poussee du meme fichier "Refonte BEL" (slide 2, sommaire) :
 * meme apres le troisieme garde-fou ci-dessus, celui-ci ne se declenchait
 * toujours jamais - AUCUNE forme du slide n'entrait jamais dans la branche
 * {@code forced}, y compris le titre du slide et les quatre items du
 * sommaire, tous retrecis sans discernement (jusqu'a 24% pour l'item le plus
 * touche, a cause du paragraphe fantome ci-dessus qui s'ajoute au probleme).
 * Cause, verifiee dans le code source d'Apache POI 5.2.5 ({@code
 * XSLFTextShape.getTextAutofit()}) : cette methode ne lit QUE le {@code
 * <a:bodyPr>} de la forme sur la SLIDE elle-meme, sans jamais remonter a la
 * mise en page ni au masque, et retombe sur {@code NORMAL} (retrecissement
 * SYSTEMATIQUE, sans verification de collision) des que rien n'y est
 * explicitement declare - motif OOXML pourtant tres courant, y compris
 * quand le masque declare bel et bien {@code <a:noAutofit/>} sur le
 * placeholder generique dont la forme herite (comme c'est le cas ici). Un
 * correctif general (traiter tout autofit non declare au niveau slide comme
 * {@code NONE}, quel que soit le fichier) toucherait potentiellement un tres
 * grand nombre de formes dans un tres grand nombre de fichiers - portee
 * jugee trop large pour etre appliquee sans discernement ; choix fait avec
 * l'utilisateur d'un correctif plus etroitement circonscrit.
 *
 * <p><b>Correctif retenu, deliberement circonscrit</b> : l'elargissement
 * (traitement special documente ci-dessous - voir {@link #fitOverflowingText})
 * n'est applique QUE lorsque les DEUX conditions suivantes sont reunies pour
 * une forme donnee :
 * <ol>
 * <li>son autofit n'est explicitement declare NI {@code noAutofit} NI
 * {@code normAutofit} NI {@code spAutoFit} dans son propre {@code
 * <a:bodyPr>} au niveau slide (voir {@link
 * #isAutofitExplicitlyDeclaredLocally}) ;</li>
 * <li>le TITRE de la diapositive correspond a un intitule de sommaire/table
 * des matieres connu, dans l'une des langues retenues pour ce projet (voir
 * {@link #TABLE_OF_CONTENTS_TITLES}) - comparaison insensible a la casse et
 * aux accents, sur le texte COMPLET et normalise du titre (voir {@link
 * #normalizeTitle}). Choix delibere d'une correspondance EXACTE plutot qu'une
 * simple inclusion de sous-chaine : un titre "Sommaire du projet XYZ" ne
 * correspondrait PAS a "Sommaire" avec ce reglage - a elargir en "contient"
 * plutot qu'"egale" si un fichier reel l'exige un jour.</li>
 * </ol>
 * Le titre de la diapositive est detecte via {@link TitleDetector} (voir
 * {@link #isSommaireSlide}), partagee avec {@link TitleRepainter} : elle
 * resout elle-meme le type de placeholder en suivant l'heritage slide -&gt;
 * mise en page -&gt; masque via {@link TitlePlaceholderResolver} - necessaire
 * pour reconnaitre un titre "centre" ({@code ctrTitle}) dont le type n'est
 * declare qu'au niveau de la mise en page, POI souffrant du meme defaut de
 * resolution d'heritage sur {@code getPlaceholder()} que sur {@code
 * getTextAutofit()} (voir la Javadoc de {@link TitlePlaceholderResolver} pour
 * le detail, verifie de la meme maniere dans le code source de POI).
 *
 * <p><b>Traitement applique aux formes concernees, revu le 2026-08-31 -
 * jamais retrecies, plutot que collision-gated</b> : une premiere version
 * appliquait a ces formes le meme traitement "collision-gated" que {@code
 * NONE} ci-dessus (retrecissement uniquement si collision reellement
 * detectee avec une autre forme de texte). Observee sur le fichier reel
 * ayant motive cet elargissement (slide 2, les quatre items du sommaire,
 * tous a la MEME taille de police heritee, aucun n'ayant de taille locale) a
 * produire une incoherence visuelle au sein de la liste : deux items
 * ("Trajectoire d'adressage" et "Demande de chiffrage") jamais retrecis (pas
 * de collision detectee - dont le dernier item de la liste, qui ne peut
 * structurellement jamais entrer en collision, n'ayant aucun voisin en
 * dessous), les deux autres retrecis a 70% (collision detectee) - un verdict
 * qui ne depend que de la geometrie propre a chaque forme (l'ecart, variable
 * d'un item a l'autre, jusqu'a son voisin immediat), sans aucun rapport avec
 * le contenu ou la longueur reelle du texte de chaque item. Confirme par
 * l'utilisateur en comparant avec le rendu PowerPoint d'origine du fichier :
 * les quatre items s'y affichent de facon uniforme, tous a taille pleine -
 * jamais retrecis. Une forme relevant de cet elargissement n'est donc plus
 * jamais retrecie, quelle que soit une eventuelle collision mesuree : elle
 * est laissee intacte a sa taille heritee (voir {@link
 * #fitOverflowingText}). Le garde-fou collision-gated original reste
 * inchange pour les formes {@code noAutofit} explicitement declarees comme
 * telles (motif {@code forcedByDeclaredNone} plus bas) : la decouverte
 * ci-dessus ne remet pas en cause son utilite pour le fichier qui l'a motive
 * (voir "Deuxieme garde-fou" plus haut).
 *
 * <p><b>Angle mort residuel (2026-08-31) et sa prise en charge</b> : une forme
 * exemptee ci-dessus, si son texte a taille pleine depasse sa PROPRE ancre
 * (motif frequent - l'ancre d'un item de sommaire est souvent dimensionnee
 * pour une seule ligne) ET chevauche par consequent une forme voisine, n'est
 * plus corrigee du tout par cette classe. Combler cet angle mort ici
 * romprait le principe retenu ci-dessus (ne plus jamais rien retrecir pour
 * ce cas) ; {@link NeighborShapeOverlapFixer} - deja concu pour reduire
 * interligne puis, en repli, police, exactement le type de correction legere
 * qu'un tel cas reclame (l'ecart constate reste faible, "quasiment pas
 * d'espace" chez PowerPoint) - est etendu pour prendre en charge CE cas
 * precis, via {@link #isSommaireBroadeningExempt} partagee entre les deux
 * classes (voir sa Javadoc et celle de {@link NeighborShapeOverlapFixer},
 * section "Formes exemptees de l'elargissement sommaire").
 *
 * <p><b>Portee assumee</b> : cette correction ne traite PAS le meme motif sur
 * des diapositives dont le titre ne correspond a aucun intitule connu - le
 * defaut structurel de POI ({@code getTextAutofit()} sans resolution
 * d'heritage) y subsiste tel quel, y compris ailleurs dans ce meme fichier.
 * Choix delibere avec l'utilisateur : portee etroite et predictible plutot
 * qu'un correctif general au rayon d'action incertain.
 *
 * <h2>Elargissement general (EXPERIMENTAL, 2026-09-05) : chaine d'heritage
 * entierement verifiee, sans dependre du titre de la diapositive</h2>
 * <p>Cas reel ayant motive cet ajout, documente en detail dans le markdown de
 * suivi du projet ({@code conversion_pptx_vers_images.md}, section 26) :
 * slide 16 du fichier "fichier-test-B.pptx",
 * deux formes cote a cote (« Espace reserve du texte 2 » et « Espace reserve
 * du contenu 3 »), toutes deux avec un {@code <a:bodyPr>} vide (aucun autofit
 * declare localement) - EXACTEMENT le motif de "l'elargissement cible aux
 * diapositives sommaire" ci-dessus, sauf que le titre de cette diapositive
 * ("Titre de diapositive utilisateur") ne
 * correspond a aucun intitule de {@link #TABLE_OF_CONTENTS_TITLES} : la forme
 * de droite, dont le contenu frole sa boite (debordement mesure d'a peine
 * ~2%, tres probablement lui-meme un artefact de la surestimation Java2D
 * documentee en tete de cette classe), se retrouve donc retrecie d'environ
 * 8% sans aucune verification de collision - confirme par log utilisateur
 * (tailles de police reellement resolues au dessin, avant/apres identiques
 * a gauche, visiblement reduites a droite).
 *
 * <p><b>Principe retenu</b> : plutot que d'elargir la liste de titres reconnus
 * (portee toujours arbitraire et jamais exhaustive), verifier directement,
 * pour la forme concernee, si le VRAI defaut structurel de POI identifie plus
 * haut ({@code getTextAutofit()} qui ne lit que le {@code bodyPr} de la
 * SLIDE) peut effectivement etre en cause - c'est-a-dire si aucun maillon de
 * la chaine d'heritage complete (slide, PUIS mise en page, PUIS masque) ne
 * declare explicitement {@code noAutofit}/{@code normAutofit}/
 * {@code spAutoFit}. {@link #isAutofitUndeclaredThroughoutInheritanceChain}
 * suit cette chaine via {@code XSLFShape#fetchShapeProperty} (la meme methode
 * publique de resolution d'heritage deja utilisee ailleurs dans ce paquetage
 * pour la meme classe de defaut - voir {@link TitlePlaceholderResolver} et
 * {@code PlaceholderGeometryResolver}, cites dans sa propre Javadoc), en
 * appliquant {@link #isAutofitExplicitlyDeclaredLocally} a CHAQUE maillon
 * visite (pas seulement la forme sur la slide). Si aucun maillon ne declare
 * quoi que ce soit, la conclusion "autofit reellement absent partout" est
 * fiable a 100% (ce n'est plus une simple absence locale pouvant cacher une
 * intention exprimee plus haut dans la chaine) - la forme est alors traitee
 * comme {@code NONE} (collision-gated), quel que soit le titre de la
 * diapositive. Ceci generalise l'elargissement "sommaire" existant sans lui
 * retirer sa propre logique (les deux exemptions restent actives en
 * parallele, voir {@link #isAutofitBroadeningExempt}).
 *
 * <p><b>Pourquoi ce n'est PAS le comportement par defaut</b> : contrairement a
 * l'elargissement "sommaire" (valide sur un fichier reel avant d'etre
 * adopte), celui-ci n'a pas encore ete confronte a l'ensemble des fichiers
 * deja corriges par ce projet - une forme aujourd'hui retrecie a raison par
 * le garde-fou {@code NORMAL}/{@code SHAPE} general (parce qu'elle deborde
 * reellement dans PowerPoint aussi, meme sans autofit déclaré nulle part -
 * cas non exclu par construction) pourrait cesser de l'etre. Expose via
 * {@link io.github.atlan77c.pptx2picture.RenderOptions.Builder#broadenAutofitExemption(boolean)},
 * desactive par defaut : a activer pour valider (slide 16, puis re-verifier
 * les fichiers "sommaire" et tout autre fichier connu de ce projet, avant
 * d'envisager de le rendre definitif et de retirer la dependance au titre).
 *
 * <h2>Quatrieme garde-fou : point de retour a la ligne "a la limite" (2026-09-02)</h2>
 * <p>Tout ce qui precede compense un ecart de mesure Java2D/PowerPoint sur l'axe
 * VERTICAL (hauteur de texte). Decouvert sur le slide 5 du fichier "fichier-test-A.pptx" ("ZoneTexte 41") : le meme type d'ecart existe aussi sur l'axe
 * HORIZONTAL, au niveau du retour a la ligne lui-meme. Confirme par comparaison
 * directe avec une capture d'ecran PowerPoint fournie par l'utilisateur : a la
 * meme taille de police, PowerPoint casse la ligne apres "avec la", tandis que
 * cette bibliotheque (Apache POI/Java2D) estime que "avec la solution Xpé" tient
 * encore sur la meme ligne - un mot de trop, qui deborde alors visuellement du
 * cadre colore en arriere-plan. Aucun run de la forme concernee ne declare de
 * police locale (heritage du theme, famille Calibri - une police standard
 * Windows) : la substitution de police pour cause de police manquante a ete
 * ecartee comme cause probable.
 *
 * <p><b>Premiere approche (2026-09-02, insuffisante - voir amendement
 * ci-dessous)</b> : plutot que de reimplementer l'algorithme de retour a la
 * ligne de POI (risque de ne pas reproduire exactement son comportement -
 * hypertexte, ponctuation, espaces insecables...), la premiere version de ce
 * garde-fou reutilisait {@code ts.getTextHeight(Graphics2D)} (deja utilisee
 * partout ailleurs dans cette classe) de facon DIFFERENTIELLE : mesure de la
 * hauteur a la largeur REELLE de l'ancre, puis a une largeur legerement
 * reduite ({@link #WIDTH_SAFETY_MARGIN}) ; une hauteur augmentee a la largeur
 * reduite etait interpretee comme un point de coupure "a la limite".
 *
 * <p><b>Amendement (2026-09-02) : angle mort sur les runs en gras</b> - cette
 * premiere approche s'est averee insuffisante sur "ZoneTexte 41" (meme
 * fichier/slide) : le retrecissement qu'elle declenchait ne faisait pas
 * disparaitre le debordement reel constate par l'utilisateur (capture d'ecran
 * a l'appui, le mot "Xpé" restait hors de la forme coloree en arriere-plan
 * apres correctif). Cause identifiee dans le XML source : le groupe de runs
 * "outil -&gt; risque a poursuivre avec la solution Xpé " (5 runs consecutifs
 * du paragraphe concerne) est integralement declare en gras ({@code b="1"}),
 * suivi d'un run non gras (la parenthese explicative) - difference de graisse
 * visible sur les captures fournies. La technique differentielle ci-dessus
 * compare la mesure de POI a ELLE-MEME (a deux largeurs proches) : si le
 * CALCUL de decoupe de POI sous-estime systematiquement la largeur du texte
 * gras (un biais interne au moteur, independant de la largeur testee - a
 * distinguer d'un ecart Java2D/PowerPoint proprement dit), les DEUX mesures
 * (100% et {@link #WIDTH_SAFETY_MARGIN}) heritent du meme biais et restent
 * egales entre elles : rien ne bascule, le garde-fou ne se declenche jamais
 * pour ce cas precis, quelle que soit la marge choisie.
 *
 * <p><b>2e tentative (2026-09-02, abandonnee - voir amendement suivant)</b> :
 * remplacement de la technique differentielle par une simulation INDEPENDANTE
 * du retour a la ligne, construite a partir des polices Java2D reelles de
 * chaque run (taille, gras, italique) plutot que de l'estimation de POI - les
 * mots repartis sur des lignes par un algorithme glouton (largeur reelle de
 * chaque mot via {@code Graphics2D.getFontMetrics(Font)}), chaque ligne ainsi
 * obtenue recopiee comme un PARAGRAPHE distinct dans une {@link XSLFTextBox}
 * jetable (meme technique que {@link VisibleTextMeasurer}), sa hauteur
 * mesuree comparee a la hauteur reelle de la forme d'origine.
 *
 * <p><b>Amendement (2e) : retour a la mesure sur le meme objet</b> - un
 * <code>mvn verify</code> reel a fait apparaitre 6 echecs sur des formes SANS
 * AUCUN rapport avec un veritable depassement horizontal (dont un texte d'un
 * seul caractere, structurellement incapable de se couper sur 2 lignes). Un
 * log de diagnostic temporaire (valeurs de hauteur brutes) a revele la cause :
 * pour un texte tenant sur une seule ligne simulee, la hauteur de la copie
 * jetable correspond EXACTEMENT a la hauteur reelle (aucun ecart) ; des que la
 * simulation repartit le texte sur PLUSIEURS lignes (donc plusieurs
 * PARAGRAPHES distincts dans la copie), un surcout constant d'environ 8-9% de
 * la hauteur simulee apparait, PAR PARAGRAPHE SUPPLEMENTAIRE, meme avec {@code
 * spaceBefore}/{@code spaceAfter} explicitement mis a 0 sur les paragraphes de
 * continuation - l'espacement entre deux PARAGRAPHES distincts n'est
 * visiblement pas rigoureusement equivalent, cote calcul de hauteur de POI, a
 * l'espacement entre deux LIGNES d'un meme paragraphe reellement coupe par
 * retour a la ligne. Comparer la hauteur d'un objet RECONSTRUIT (plusieurs
 * paragraphes synthetiques) a la hauteur de l'objet D'ORIGINE (un seul
 * paragraphe reellement coupe) introduit donc un biais systematique, croissant
 * avec le nombre de lignes simulees - independant de tout veritable
 * depassement de largeur, et suffisant a lui seul pour declencher un
 * retrecissement injustifie des qu'un texte de plusieurs mots occupe deja 2
 * lignes ou plus, gras ou non.
 *
 * <p><b>Correction retenue</b> : abandon de toute reconstruction d'objet
 * jetable pour ce garde-fou. Retour a la technique differentielle du tout
 * premier essai ({@link #isWrapMarginUnstable}, mesures {@code
 * ts.getTextHeight(Graphics2D)} sur l'ancre REELLE de {@code ts}, mutee
 * temporairement puis aussitot restauree - comme {@link
 * #computeOverflowZones} et le reste de cette classe, deux mesures sur le
 * MEME objet ne peuvent par construction jamais diverger a cause d'un modele
 * de hauteur different), mais avec une largeur reduite qui n'est plus une
 * marge fixe seule : {@link #maxBoldWidthPremium} mesure, via {@code
 * Graphics2D.getFontMetrics(Font)}, l'ecart entre la largeur REELLE (en gras)
 * et la largeur qu'aurait le meme mot SANS le gras, pour le mot le plus
 * penalisant parmi tous les runs en gras de la forme - cible directement
 * l'hypothese diagnostiquee sur "ZoneTexte 41" (sous-estimation du gras par le
 * calcul de decoupe de POI), ajoutee a {@link #WIDTH_SAFETY_MARGIN} pour la
 * reduction de largeur testee. Pour une forme sans aucun run en gras, cet
 * ajout vaut 0 : le comportement redevient alors EXACTEMENT celui de la marge
 * generique seule, c'est-a-dire celui, deja eprouve, d'avant ce garde-fou
 * horizontal - aucun risque de regression sur les formes non concernees par
 * cette hypothese specifique.
 *
 * <p>Levier de correction inchange (retrecissement de police par paliers de
 * {@link #STEP}). Portee, placement dans {@link #fitOverflowingText} (en tout
 * premier, avant le reste du traitement de la forme, y compris avant la
 * sortie anticipee "forced sans collision verticale reelle" - voir plus bas)
 * et logique de repli en cas d'echec (restauration si aucun retrecissement ne
 * suffit jusqu'a {@link #MIN_SCALE}) inchanges par ailleurs.
 *
 * <p><b>Portee assumee</b> : {@link #maxBoldWidthPremium} ne mesure que les
 * runs EN GRAS dont la taille de police est declaree localement (secours 0,
 * sans effet sur la marge, sinon) ; une famille de police non declaree
 * localement (heritage du theme) retombe sur {@link #DEFAULT_FONT_FAMILY}
 * (Calibri - police par defaut du theme Office standard) faute de resolution
 * d'heritage complete ici, pour la mesure de largeur gras/non-gras
 * uniquement ; le mot le plus penalisant PARMI TOUS LES RUNS EN GRAS de la
 * forme est utilise comme majorant global (pas necessairement celui de la
 * ligne effectivement a la limite) - approximation deliberement simple et
 * prudente (majorant, jamais sous-estime) plutot qu'une identification
 * precise de la ligne concernee. Pour un texte sans aucun gras, ce garde-fou
 * se comporte exactement comme la marge generique {@link #WIDTH_SAFETY_MARGIN}
 * seule. A affiner si un autre fichier reel montre une limite de cette
 * approche.
 *
 * <h2>Piste explorée puis abandonnée : marge de tolerance sur la detection de
 * collision verticale (2026-09-02, retiree le meme jour)</h2>
 * <p>Sur le slide 5 du fichier "fichier-test-A.pptx", forme "ZoneTexte 39" :
 * le texte ("texte utilisateur (segments courts)") deborde
 * bel et bien de son anchor selon la mesure de cette bibliotheque (jusqu'a
 * 203,29pt de texte mesure contre 84,36pt d'anchor - voir le log de
 * diagnostic "debordement mesure... mais aucune collision"), et un premier
 * rendu de l'utilisateur semblait montrer "Usage" chevauchant visuellement
 * l'en-tete "USAGERS" (forme "ZoneTexte 37"). Une piste a ete tentee : elargir
 * chaque zone de debordement d'une marge fixe en points ({@code
 * COLLISION_MARGIN_POINTS}, via un nouveau {@code growForCollisionTest})
 * avant de tester le chevauchement avec les formes voisines dans {@link
 * #findCollidingShape} (et donc {@link #overflowCollidesWithText}) -
 * {@link #computeOverflowZones} elle-meme restant inchangee (contrat
 * geometrique exact intact, y compris pour ses propres tests unitaires).
 * Une premiere valeur (6pt, a partir d'une analyse XML manuelle laissant
 * supposer un ecart de ~4,95pt) s'est averee insuffisante en pratique (le
 * chevauchement persistait sur le rendu reel) ; reprise precise du calcul
 * geometrique a partir du XML source et des valeurs exactes du log a montre
 * que l'ecart reel etait en fait d'environ 6,87pt, et la marge a ete portee a
 * 8pt en consequence.
 *
 * <p><b>Abandon : le cas motivant etait un FAUX POSITIF</b> - avec la marge a
 * 8pt, le texte de "ZoneTexte 39" s'est retrouve fortement retreci... mais
 * l'utilisateur, en verifiant directement le fichier source dans PowerPoint,
 * a confirme qu'AUCUN chevauchement visuel n'existe reellement entre
 * "ZoneTexte 39" et "USAGERS" - taille de police d'origine incluse. Le
 * probleme n'a donc jamais ete un veritable "quasi-echec" de detection de
 * collision : la mesure de depart elle-meme (203,29pt necessaires selon
 * cette bibliotheque) est fortement surestimee par rapport a ce que
 * PowerPoint affiche reellement pour ce meme contenu. Aucune marge de
 * tolerance, quelle que soit sa valeur, ne peut corriger une DETECTION de
 * collision fondee sur une mesure de depart erronee - elle ne fait que
 * deplacer le seuil auquel l'erreur devient visible (marge trop faible :
 * collision reelle manquee comme avant tout ce garde-fou ; marge trop forte,
 * comme ici : retrecissement injustifie d'un texte qui n'en avait pas
 * besoin). Hypothese envisagee pour la surestimation (fonte "PoliceX",
 * police d'Etat francaise utilisee par cette forme, potentiellement resolue
 * differemment par Java2D que par PowerPoint) - ECARTEE : l'utilisateur a
 * confirme que "PoliceX" est bien installee sur sa machine. La cause exacte
 * de la surestimation reste donc non identifiee a ce stade (a rapprocher de
 * l'ecart Java2D/PowerPoint deja documente ailleurs dans cette classe,
 * jusqu'a 30-35% observe sur d'autres fichiers - voir {@link
 * #SAFETY_MARGIN}) ; a reprendre uniquement si un futur fichier reel fournit
 * de nouvelles donnees exploitables (idealement une comparaison directe des
 * metriques de police Java2D vs PowerPoint pour "PoliceX"), plutot que de
 * continuer a ajuster une marge de detection qui ne peut pas compenser une
 * mesure de hauteur erronee a la source.
 *
 * <p><b>Retenu</b> : {@link #findCollidingShape} (et {@link
 * #overflowCollidesWithText}) sont revenues a la detection de collision
 * geometrique STRICTE, sans marge de tolerance - le comportement d'origine,
 * anterieur a cette piste, qui ne declenchait jamais de retrecissement pour
 * "ZoneTexte 39" et correspond donc au rendu reel confirme par l'utilisateur.
 *
 * <h2>Cinquieme garde-fou : fidelite de l'interligne en pourcentage (spcPct)
 * (2026-09-02)</h2>
 * <p>Suite a la piste ci-dessus (abandonnee), l'utilisateur a reformule le vrai
 * probleme : ce n'est pas une collision a corriger, c'est un ecart de FIDELITE
 * VISUELLE - l'espace ENTRE LES LIGNES a l'interieur de "ZoneTexte 39" est
 * nettement plus grand dans le rendu de cette bibliotheque que dans PowerPoint,
 * ce qui gonfle la hauteur totale de la forme (et c'est CA qui causait le
 * "chevauchement" visuel du tout premier rendu, avant meme le detecteur de
 * collision). Citation exacte : <i>"Il ne faut pas corriger le chevauchement
 * mais reduire ou reproduire l'espace entre les lignes d'une meme zone de texte
 * pour rester fidele au powerpoint."</i>
 *
 * <p><b>Cause racine, prouvee par le code source reel d'Apache POI 5.2.5</b>
 * (recupere et lu directement via {@code raw.githubusercontent.com/apache/poi/
 * REL_5_2_5/...} - maven central bloque depuis l'environnement d'investigation,
 * GitHub accessible) : {@code org.apache.poi.sl.draw.DrawTextParagraph#draw}
 * calcule l'avance verticale entre deux lignes comme {@code (spacing*0.01) *
 * line.getHeight()}, ou {@code line.getHeight()} ({@code DrawTextFragment}) vaut
 * {@code ascent+descent+leading} - {@code leading} etant {@code
 * TextLayout.getLeading()} SAUF si cette valeur vaut 0 (frequent pour de
 * nombreuses polices), auquel cas POI substitue {@code (ascent+descent)*0.15},
 * avec ce commentaire explicite dans le code source officiel : <i>"we use a
 * 115% value instead of the 120% proposed one, as this seems to be closer to
 * LO/OO"</i> (LibreOffice/OpenOffice - pas PowerPoint). Confirme en environnement
 * reel par un diagnostic temporaire (desormais retire) : pour tous les runs de
 * "ZoneTexte 39" (police "PoliceX", 17.64pt), {@code TextLayout.getLeading()}
 * vaut bien 0.0 - le repli +15% de POI est bien emprunte ici.
 *
 * <p><b>Calibration reelle (PowerPoint 365, confirme par capture d'ecran du
 * ruban/panneau lateral)</b> : en dupliquant "ZoneTexte 39" et en activant
 * temporairement "Ajuster la forme au texte" ({@code <a:spAutoFit/>}) sur des
 * variantes controlees du contenu reel (fichier de test dedie {@code
 * C:\temp\Lib_pptx_picture\2026092-test.pptx}), trois experiences ont ete
 * decisives : (1) passer le texte en gras ne change PAS la hauteur calculee par
 * PowerPoint ; (2) changer de police (PoliceX vers Times New Roman, meme
 * taille/contenu) ne la change PAS non plus - exactement la meme hauteur dans
 * les deux cas ; (3) a taille de police fixee, la hauteur reste coherente avec
 * un modele simple sur deux tailles tres differentes (17.64pt et 32pt, ce
 * dernier avec un nombre de lignes different du fait du retour a la ligne).
 * Conclusion : contrairement au modele de POI (ascent/descent REELS, sensibles
 * a la police et a la graisse), PowerPoint calcule la hauteur d'une ligne simple
 * de facon INDEPENDANTE de la police et de la graisse - uniquement a partir de
 * la taille de police DECLAREE, selon {@code hauteur_ligne = taille x
 * LINE_HEIGHT_FIDELITY_FACTOR}. Deux mesures reelles independantes donnent
 * ~1.2154 (17.64pt) et ~1.1983 (32pt) - moins de 1.5% d'ecart sur une taille
 * quasiment doublee - d'ou la valeur retenue de {@link
 * #LINE_HEIGHT_FIDELITY_FACTOR} (1.2, tres proche des deux mesures et de la
 * convention DTP usuelle "interligne simple ~= 1.2x la taille de police").
 *
 * <p><b>Correction retenue</b> : {@link #correctPercentLineSpacingForFidelity}
 * ne reimplemente pas la mise en page hors de POI - elle mute in-place, pour
 * chaque paragraphe a interligne en pourcentage dont le run dominant resout un
 * leading Java2D de 0, la valeur d'interligne DECLAREE ({@code
 * XSLFTextParagraph#setLineSpacing}) vers une valeur corrigee qui compense
 * exactement le repli +15% de POI, de sorte que le dessin natif de POI ({@code
 * slide.draw(graphics)}, inchange par ailleurs) produise un resultat visuel
 * fidele a PowerPoint - et, en consequence directe, que {@code
 * ts.getTextHeight(graphics)} (utilise par tout le reste de cette classe) soit
 * lui aussi fidele, sans calcul de hauteur separe. Appelee en tout premier dans
 * {@link #fitOverflowingText}, pour toute forme non exemptee (sommaire/table des
 * matieres), que l'autofit soit force ou non - c'est un defaut de rendu visuel,
 * independant de la logique de collision des autres garde-fous.
 *
 * <p><b>Portee assumee, limites connues</b> (voir Javadoc de la methode pour le
 * detail) : se cale sur le run de plus grande taille du paragraphe pour les
 * paragraphes a tailles mixtes (approximation) ; ne s'applique qu'aux runs dont
 * le leading Java2D resout a 0 (une police avec son propre leading natif n'est
 * pas connue pour souffrir du meme ecart, faute de donnee reelle sur ce cas) ;
 * ne corrige pas la hauteur propre de la toute premiere ligne du bloc (jamais
 * multipliee par le %% d'interligne dans l'algorithme de POI, quelle que soit la
 * valeur fournie) - erreur residuelle marginale sur un bloc de plusieurs lignes,
 * plus sensible sur un bloc d'une seule ligne. Calibration fondee sur deux
 * mesures reelles (une seule police testee en detail, "PoliceX", et Times New
 * Roman en confirmation ponctuelle) : a reprendre si un futur fichier reel
 * montre un ecart significatif avec une police tres differente ou un interligne
 * autre que "Multiple" (spcPct).
 *
 * <h2>Exemption des bulles narratives (callouts) de tout chevauchement (2026-09-03)</h2>
 * <p>Suite a une regression signalee sur le slide 4 ("Title 3" retreci a une
 * taille de police illisible) apres le deploiement du cinquieme garde-fou
 * ci-dessus : le journal de debogage reel a montre que {@link
 * NeighborShapeOverlapFixer} detectait un chevauchement entre "Title 3" et
 * "Bulle narrative : rectangle a coins arrondis 1" (une bulle de rappel
 * pointant, via un bec fin, vers une ligne precise du texte de "Title 3") et
 * tentait de le resoudre en cumulant reduction d'interligne puis, en repli,
 * reduction de police jusqu'a 24% de la taille d'origine. Une capture d'ecran
 * de PowerPoint reel a confirme que le texte de la bulle et celui de "Title 3"
 * ne se chevauchent JAMAIS visuellement - seul le bec (fin trait de connecteur)
 * de la bulle traverse la zone du texte. Or {@link
 * NeighborShapeOverlapFixer#findCollidingShape} (comme {@link
 * #findCollidingShape} ci-dessous) compare a l'ANCRE ENTIERE de la forme
 * voisine (limite deja documentee dans leurs Javadoc respectives) - et l'ancre
 * OOXML d'une forme de type "callout" (bulle de rappel) englobe geometriquement
 * le trace complet de la forme, bec compris, jusqu'a l'endroit qu'il designe.
 * Le "chevauchement" mesure ici est donc un artefact de cette limite de mesure
 * (comparaison a la boite englobante plutot qu'au contenu visible), pas un
 * veritable chevauchement de texte.
 *
 * <p><b>Confirme par l'utilisateur</b> : dans le fichier source, ces bulles
 * narratives sont deliberement replacees au premier plan (ordre Z) a la toute
 * fin de la composition, precisement pour pointer par-dessus d'autres formes -
 * un chevauchement avec elles est donc TOUJOURS intentionnel et ne doit jamais
 * declencher de correction, ni pour la bulle elle-meme (qui n'est jamais
 * retrecie/interlignee a cause d'une forme voisine) ni pour la forme qu'elle
 * survole (jamais retrecie a cause d'une bulle).
 *
 * <p><b>Correction retenue</b> : toute forme dont le type de geometrie
 * PowerPoint ({@link XSLFTextShape#getShapeType()}) appartient a la famille
 * "callout" de {@link ShapeType} (detectee via {@link #isCalloutShape}, qui
 * teste si {@code ShapeType.name()} contient {@code "CALLOUT"} - couvre {@code
 * CALLOUT_1..3}, les variantes "accent"/"border", les fleches de rappel, {@code
 * WEDGE_RECT_CALLOUT}, {@code WEDGE_ROUND_RECT_CALLOUT} (le type exact observe
 * ici, nomme par defaut "Bulle narrative : rectangle a coins arrondis" par
 * PowerPoint en francais), {@code WEDGE_ELLIPSE_CALLOUT}, {@code
 * CLOUD_CALLOUT}, etc. - plutot qu'une correspondance sur le seul type observe,
 * pour couvrir sans avoir a y revenir toute autre forme de bulle utilisee dans
 * ce meme but) est desormais exclue des DEUX cotes de toute detection de
 * chevauchement ENTRE FORMES DISTINCTES : {@link #findCollidingShape} (ce
 * fichier) et {@link NeighborShapeOverlapFixer#findCollidingShape} ne la
 * retiennent plus jamais comme forme "colliding", et {@link
 * NeighborShapeOverlapFixer#fixNeighborOverlaps} l'ignore entierement en tout
 * debut de boucle (jamais elle-meme corrigee a cause d'une collision avec une
 * forme voisine). Le debordement d'une bulle dans SA PROPRE ancre (motif sans
 * rapport avec une forme voisine, garde-fous habituels de {@link
 * #fitOverflowingText}) n'est en revanche pas concerne par cette exemption et
 * reste traite normalement.
 *
 * <p><b>Portee assumee</b> : {@link #isCalloutShape} est partagee
 * (package-private) entre cette classe et {@link NeighborShapeOverlapFixer},
 * meme convention que {@link #isSommaireBroadeningExempt}. Choix delibere par
 * l'utilisateur d'une exemption totale et definitive, plutot qu'une resolution
 * plus fine (par exemple en excluant uniquement le bec de l'ancre de la bulle,
 * information non exposee par l'API POI sans analyse du XML de la forme).
 *
 * <h2>Hors de portee : les cellules de tableau natif (2026-09-04)</h2>
 * <p>Cette classe (comme {@link NeighborShapeOverlapFixer}) ne traite que les formes
 * retournees par {@link #collectTextShapes}, qui ne descend que dans les groupes ({@link
 * XSLFGroupShape}) - jamais dans un {@code XSLFTable}. Un tableau n'est ni un groupe ni
 * une {@link XSLFTextShape} au sens de ce parcours, donc ni lui ni ses cellules (qui
 * pourtant EN SONT une - {@code XSLFTableCell extends XSLFTextShape}) ne beneficient
 * d'aucun garde-fou de cette classe, y compris le cinquieme (fidelite d'interligne)
 * ci-dessus, alors meme que POI fait grandir dynamiquement la hauteur de chaque ligne de
 * tableau selon le texte mesure de ses cellules ({@code XSLFTable#updateCellAnchor()},
 * pass #1 : {@code rowHeights[row] = Math.max(rowHeights[row], maxHeight)} avec {@code
 * maxHeight} issu de {@code DrawTextShape#getTextHeight()}) - exactement le meme
 * mecanisme de surestimation Java2D/AWT que le cinquieme garde-fou corrige pour les
 * formes normales. Voir {@link TableCellLineSpacingFixer}, qui reutilise {@link
 * #correctPercentLineSpacingForFidelity} (partagee, package-private) pour combler cette
 * lacune specifiquement pour les cellules de tableau, avec son propre retrecissement de
 * secours (reutilisant {@link #STEP}, {@link #MIN_SCALE}, {@link #MAX_ITER}, {@link
 * #SAFETY_MARGIN} et {@link #captureBaselineFontSizes}, egalement partages) plutot que
 * d'etendre {@link #fitOverflowingText} lui-meme : la plupart de ses autres garde-fous
 * (marge de retour a la ligne, exemption sommaire/callout, mesure des paragraphes vides
 * en fin de bloc...) sont calibres et testes uniquement pour des formes autonomes, jamais
 * pour la geometrie particuliere d'un tableau (lignes qui grandissent dynamiquement,
 * cellules fusionnees) - les y exposer sans calibration dediee aurait ete un risque de
 * regression non maitrise sur l'ensemble des fichiers deja confirmes par l'utilisateur.
 */
public final class OverflowAwareTextFitter {

    private static final Logger LOG = LoggerFactory.getLogger(OverflowAwareTextFitter.class);

    /**
     * Pas de reduction applique a chaque iteration (mimique les paliers utilises par
     * PowerPoint). Visibilite package (pas {@code private}) : reutilise telle quelle par
     * {@link TableCellLineSpacingFixer} pour son propre retrecissement de secours, plutot
     * que de dupliquer une valeur calibree - voir Javadoc de cette classe.
     */
    static final double STEP = 0.02;
    static final double MIN_SCALE = 0.25;
    static final int MAX_ITER = 38;

    /**
     * Marge de securite appliquee a la hauteur cible lors du retrecissement :
     * on arrete de retrecir des que le texte mesure tient dans {@code anchor.getHeight() * SAFETY_MARGIN}
     * plutot que dans {@code anchor.getHeight()} strictement. Sans cette marge, un cas reel
     * (diagnostic sur un fichier de production - forme "spAutoFit" a plusieurs paragraphes, ecart de
     * mesure de ~30% entre Java2D et PowerPoint, dans le haut de la fourchette deja documentee)
     * a montre que le retrecissement s'arretait des que getTextHeight() repassait sous
     * anchor.getHeight(), mais avec une marge residuelle de moins de 1pt sur une hauteur de plusieurs
     * dizaines de points : largement insuffisant pour absorber a la fois l'epaisseur du trait de
     * bordure de la forme et un eventuel ecart residuel entre la mesure de getTextHeight() et ce que
     * slide.draw() peint reellement - le texte continuait donc a chevaucher visuellement la bordure
     * malgre un retrecissement de police tres marque (jusqu'a -30% environ dans ce cas).
     * Visibilite package (pas {@code private}) : reutilisee telle quelle par {@link
     * TableCellLineSpacingFixer}, meme raison que {@link #STEP}.
     */
    static final double SAFETY_MARGIN = 0.97;

    /**
     * Marge de securite appliquee a la largeur reduite testee lors de la
     * mesure differentielle sur le meme objet (voir {@link
     * #isWrapMarginUnstable} et Javadoc de la classe, section
     * "Quatrieme garde-fou..."). Meme ordre de grandeur que {@link
     * #SAFETY_MARGIN}, faute de mesure plus precise de l'ecart reel
     * Java2D/PowerPoint sur cet axe - a affiner si un fichier reel montre
     * qu'elle est insuffisante ou trop large.
     */
    private static final double WIDTH_SAFETY_MARGIN = 0.97;

    /**
     * Famille de police utilisee pour resoudre la largeur reelle d'un mot
     * gras (voir {@link #maxBoldWidthPremium}) quand un run ne declare
     * aucune famille localement (heritage du theme, non resolu ici - voir
     * Javadoc de la classe, section "Quatrieme garde-fou...", "Portee
     * assumee"). Calibri est la police par defaut du theme Office standard,
     * deja identifiee comme la police effective de la forme ayant motive ce
     * garde-fou.
     */
    private static final String DEFAULT_FONT_FAMILY = "Calibri";

    /**
     * Facteur de conversion "taille de police -> hauteur de ligne simple" utilise
     * par PowerPoint pour un interligne "Multiple" (spcPct), etabli par
     * calibration REELLE (2026-09-02, voir Javadoc de la classe, section
     * "Cinquieme garde-fou..." pour le recit complet de l'investigation) :
     * contrairement au modele d'Apache POI (ascent+descent+leading mesures via
     * Java2D/{@code TextLayout}, dependants de la police et de la graisse),
     * PowerPoint calcule la hauteur d'une ligne simple comme {@code taille_police
     * x LINE_HEIGHT_FIDELITY_FACTOR}, INDEPENDAMMENT de la police et de la graisse
     * - confirme experimentalement sur le fichier reel de l'utilisateur : gras
     * active/desactive et police PoliceX/Times New Roman donnent tous la MEME
     * hauteur PowerPoint pour un meme contenu/taille (aucun effet mesurable).
     * Valeur calibree a partir de deux mesures reelles independantes dans
     * PowerPoint 365 (via "Ajuster la forme au texte" sur une copie de la forme
     * reelle du fichier de l'utilisateur) : ~1.2154 a 17.64pt (6 lignes) et
     * ~1.1983 a 32pt (10 lignes, retour a la ligne different de par la taille) -
     * moins de 1.5% d'ecart sur une taille quasiment doublee. 1.2 retenu : valeur
     * ronde, tres proche des deux mesures (mieux alignee sur celle a 32pt), et
     * correspond a la convention DTP usuelle "interligne simple ~= 1.2x la taille
     * de police".
     */
    private static final double LINE_HEIGHT_FIDELITY_FACTOR = 1.2;

    /**
     * Intitules de sommaire/table des matieres reconnus, dans les langues
     * retenues pour ce projet (voir {@link TitleDetector#TITLE_NAME_PREFIXES},
     * meme jeu de langues) - deja normalises (voir {@link #normalizeTitle}) au
     * chargement de la classe. Voir Javadoc de la classe, section
     * "Elargissement cible..." pour l'usage et le choix d'une correspondance
     * exacte plutot qu'une inclusion de sous-chaine.
     */
    private static final Set<String> TABLE_OF_CONTENTS_TITLES = buildNormalizedTitleSet(
            "Sommaire", "Table des matières",              // francais
            "Contents", "Table of Contents", "Agenda",       // anglais
            "Índice", "Tabla de contenido",                  // espagnol
            "Índice", "Sumário",                             // portugais (bresilien)
            "Sommario", "Indice",                            // italien
            "Inhalt", "Inhaltsverzeichnis"                   // allemand
    );

    private OverflowAwareTextFitter() {
    }

    /**
     * Retrecit, in-place, les formes de texte du slide qui debordent
     * reellement de leur boite (voir Javadoc de la classe). Modifie les
     * tailles de police des {@link XSLFTextRun} concernes ; a appeler avant
     * {@code slide.draw(graphics)}.
     *
     * @return le nombre de formes effectivement retrecies.
     */
    public static int fitOverflowingText(XSLFSlide slide, Graphics2D graphics) {
        return fitOverflowingText(slide, graphics, false);
    }

    /**
     * Variante de {@link #fitOverflowingText(XSLFSlide, Graphics2D)} exposant
     * l'elargissement EXPERIMENTAL documente dans la Javadoc de la classe,
     * section "Elargissement general (experimental)" - voir {@code
     * broadenAutofitExemption} sur {@link
     * io.github.atlan77c.pptx2picture.RenderOptions.Builder#broadenAutofitExemption(boolean)}
     * pour son usage recommande (a activer au cas par cas pour valider avant
     * adoption definitive).
     *
     * @return le nombre de formes effectivement retrecies.
     */
    public static int fitOverflowingText(XSLFSlide slide, Graphics2D graphics, boolean broadenAutofitExemption) {
        int count = 0;
        List<XSLFShape> allTextShapes = collectTextShapes(slide.getShapes());
        // Voir Javadoc de la classe, section "Elargissement cible..." : calcule une
        // seule fois par slide, avant la boucle sur les formes.
        boolean sommaireSlide = isSommaireSlide(slide);

        for (XSLFShape shape : allTextShapes) {
            XSLFTextShape ts = (XSLFTextShape) shape;
            TextShape.TextAutofit autofit = ts.getTextAutofit();
            // NORMAL et SHAPE partagent la meme logique : PowerPoint ne montre jamais
            // de debordement reel pour ces deux modes, donc tout debordement mesure ici
            // est un artefact du calcul de metriques Java2D (jamais une intention de
            // l'auteur) et est systematiquement corrige. NONE est le cas normal ou un
            // debordement peut etre volontaire (voir "forced" ci-dessous).
            boolean forcedByDeclaredNone = autofit == TextShape.TextAutofit.NONE;

            // Elargissement cible aux diapositives sommaire/table des matieres (voir
            // Javadoc de la classe, section "Elargissement cible...", traitement revu le
            // 2026-08-31) : une forme dont l'autofit n'est pas declare localement, sur une
            // slide de ce type, n'est PLUS jamais retrecie ICI - contrairement au
            // traitement "collision-gated" de {@code forcedByDeclaredNone} ci-dessus, qui
            // produisait une incoherence au sein d'une meme liste a police heritee
            // identique (voir Javadoc de la classe pour le detail). Sortie anticipee de la
            // boucle : cette forme n'est touchee par AUCUNE des etapes suivantes de cette
            // methode (mesure, detection de collision, retrecissement) - mais reste prise
            // en charge par {@link NeighborShapeOverlapFixer}, qui partage {@link
            // #isSommaireBroadeningExempt} pour savoir qu'il doit, pour cette meme forme,
            // agir meme quand son texte depasse sa propre ancre (voir sa Javadoc).
            if (isAutofitBroadeningExempt(ts, sommaireSlide, broadenAutofitExemption)) {
                if (LOG.isDebugEnabled()) {
                    LOG.debug("{} : autofit non declare nulle part dans la chaine d'heritage (slide sommaire, ou "
                            + "elargissement general experimental actif) -> jamais retrecie ici (voir Javadoc de "
                            + "la classe ; prise en charge deleguee a NeighborShapeOverlapFixer)",
                            shape.getShapeName());
                }
                continue;
            }

            boolean forced = forcedByDeclaredNone;

            Rectangle2D anchor = ts.getAnchor();
            if (anchor == null || anchor.getHeight() <= 0) {
                continue;
            }

            // Cinquieme garde-fou : fidelite de l'interligne en pourcentage (voir Javadoc de
            // la classe, meme section, et {@link #LINE_HEIGHT_FIDELITY_FACTOR}). Volontairement
            // place ICI, avant toute mesure de hauteur (garde-fous un a quatre plus bas) : cette
            // methode mute in-place l'interligne DECLARE des paragraphes concernes pour
            // compenser le repli interne de POI, de sorte que {@code ts.getTextHeight(graphics)}
            // - utilise par tout le reste de cette methode - reflete deja une mesure fidele,
            // sans avoir besoin d'un calcul de hauteur separe. S'applique a toute forme non
            // exemptee plus haut (sommaire/table des matieres), que l'autofit soit force ou non
            // : c'est un defaut de rendu visuel, independant de la logique de collision.
            correctPercentLineSpacingForFidelity(ts, graphics);

            // Quatrieme garde-fou : point de retour a la ligne "a la limite" (voir Javadoc
            // de la classe, meme section) - axe HORIZONTAL, independant de tout ce qui suit
            // dans cette methode (axe vertical, garde-fous un a trois). Volontairement place
            // ICI, avant le reste du traitement de cette forme (notamment avant la sortie
            // anticipee "forced sans collision reelle" plus bas) : contrairement au
            // debordement vertical, dont l'absence de correction pour ce cas precis est un
            // choix delibere (voir Javadoc de la classe, "Deuxieme garde-fou"), rien ne
            // justifie qu'un depassement horizontal de la PROPRE ancre de la forme - un
            // defaut de mesure Java2D independant de toute notion de collision avec une
            // forme voisine - passe inapercu simplement parce que cette forme n'a par
            // ailleurs aucun probleme vertical. S'applique a toute forme non exemptee plus
            // haut (sommaire), quel que soit son autofit, et seulement si le retour a la
            // ligne est actif.
            if (ts.getWordWrap() && isWrapMarginUnstable(ts, anchor, graphics)) {
                Map<XSLFTextRun, Double> wrapBaseline = captureBaselineFontSizes(ts);
                if (!wrapBaseline.isEmpty()) {
                    double wFactor = 1.0;
                    int wIter = 0;
                    boolean unstable = true;
                    while (unstable && wFactor > MIN_SCALE && wIter < MAX_ITER) {
                        wFactor -= STEP;
                        for (Map.Entry<XSLFTextRun, Double> e : wrapBaseline.entrySet()) {
                            e.getKey().setFontSize(Math.max(1.0, e.getValue() * wFactor));
                        }
                        unstable = isWrapMarginUnstable(ts, anchor, graphics);
                        wIter++;
                    }
                    if (unstable) {
                        // Aucun retrecissement (jusqu'a MIN_SCALE) n'a suffi a stabiliser le
                        // point de coupure : taille restauree plutot que de produire un texte
                        // ecrase pour un gain incertain (meme philosophie que plus haut).
                        for (Map.Entry<XSLFTextRun, Double> e : wrapBaseline.entrySet()) {
                            e.getKey().setFontSize(e.getValue());
                        }
                        if (LOG.isDebugEnabled()) {
                            LOG.debug("{} : point de retour a la ligne toujours a la limite meme a la taille "
                                    + "de police minimale - taille restauree", shape.getShapeName());
                        }
                    } else {
                        // Compte independamment du reste de cette methode : un tres leger
                        // risque de compter deux fois la meme forme si elle a par ailleurs
                        // aussi besoin d'un retrecissement vertical plus bas est accepte -
                        // `count` n'est qu'un total informatif pour les logs (voir
                        // PptxSlideRenderer), jamais utilise pour une decision de correction.
                        count++;
                        if (LOG.isDebugEnabled()) {
                            LOG.debug("{} : point de retour a la ligne retreci a {}% pour eloigner le bord droit "
                                    + "du texte de la limite de l'ancre (ecart Java2D/PowerPoint possible pres "
                                    + "du bord, voir Javadoc de la classe)",
                                    shape.getShapeName(), Math.round(wFactor * 100));
                        }
                    }
                }
            }

            // Troisieme garde-fou (NONE) : un paragraphe entierement vide EN FIN DE BLOC
            // n'a aucun contenu visible mais gonfle ts.getTextHeight(graphics) comme
            // s'il s'agissait d'une ligne reelle - voir Javadoc de la classe (et son
            // amendement du 2026-09-01 : un paragraphe vide EN MILIEU de bloc reste, lui,
            // inclus dans la mesure, voir hasTrailingBlankParagraph/VisibleTextMeasurer).
            // Ne s'applique qu'aux formes noAutofit explicitement declarees comme telles
            // (forcedByDeclaredNone - les formes relevant de l'elargissement sommaire ne
            // passent plus par ici, voir la sortie anticipee ci-dessus) dont le DERNIER
            // paragraphe est entierement vide ; sans effet (mesure inchangee) dans tous
            // les autres cas.
            boolean ignoreBlankParagraphs = forced && hasTrailingBlankParagraph(ts);
            VisibleTextMeasurer visibleMeasurer = ignoreBlankParagraphs
                    ? new VisibleTextMeasurer(slide, ts, anchor, graphics) : null;
            try {
                if (forced) {
                    double maxDeclaredFontSize = maxDeclaredFontSize(ts);
                    if (maxDeclaredFontSize > 0 && maxDeclaredFontSize > anchor.getHeight()) {
                        if (LOG.isDebugEnabled()) {
                            LOG.debug("{} : taille de police declaree ({}pt) > hauteur de la boite ({}pt) -> "
                                    + "debordement volontaire de l'auteur (ex. gros caractere decoratif), non retrecie",
                                    shape.getShapeName(), maxDeclaredFontSize, anchor.getHeight());
                        }
                        continue;
                    }

                    double initialTextHeight = visibleMeasurer != null ? visibleMeasurer.height() : ts.getTextHeight(graphics);
                    VerticalAlignment valign = ts.getVerticalAlignment();
                    List<Rectangle2D> overflowZones = computeOverflowZones(anchor, initialTextHeight, valign);
                    XSLFShape collidingShape = overflowZones.isEmpty() ? null : findCollidingShape(overflowZones, ts, allTextShapes);
                    if (collidingShape == null) {
                        if (!overflowZones.isEmpty() && LOG.isDebugEnabled()) {
                            LOG.debug("{} : debordement mesure ({} > {}) mais aucune collision avec une autre "
                                    + "forme de texte -> non retrecie", shape.getShapeName(), initialTextHeight, anchor.getHeight());
                        }
                        continue;
                    }
                    if (LOG.isDebugEnabled()) {
                        LOG.debug("{} : debordement mesure ({} > {}, alignement {}) entre en collision avec '{}' -> retrecissement force",
                                shape.getShapeName(), initialTextHeight, anchor.getHeight(), valign, collidingShape.getShapeName());
                    }
                }

                Map<XSLFTextRun, Double> baseline = captureBaselineFontSizes(ts);
                if (baseline.isEmpty()) {
                    continue; // tailles heritees du theme/layout, non modifiables ici
                }

                double factor = 1.0;
                int iter = 0;
                double textHeight = visibleMeasurer != null ? visibleMeasurer.height() : ts.getTextHeight(graphics);
                boolean didShrink = false;
                double targetHeight = anchor.getHeight() * SAFETY_MARGIN;

                while (textHeight > targetHeight && factor > MIN_SCALE && iter < MAX_ITER) {
                    factor -= STEP;
                    for (Map.Entry<XSLFTextRun, Double> e : baseline.entrySet()) {
                        e.getKey().setFontSize(Math.max(1.0, e.getValue() * factor));
                    }
                    textHeight = visibleMeasurer != null ? visibleMeasurer.height() : ts.getTextHeight(graphics);
                    iter++;
                    didShrink = true;
                }

                if (forced && textHeight > targetHeight) {
                    // Le retrecissement visant un ajustement complet dans la boite n'a pas
                    // converge, meme en ecrasant la police jusqu'a sa limite basse (MIN_SCALE)
                    // ou le nombre max d'iterations (MAX_ITER) : la boite est structurellement
                    // trop petite pour ce texte (ex. legende/annotation flottante avec bien plus
                    // de texte que la boite ne pourra jamais en contenir - meme dans
                    // PowerPoint). Deuxieme passe avec un objectif different : PowerPoint
                    // n'exige jamais qu'un texte noAutofit tienne entierement dans sa boite -
                    // seulement qu'il ne produise pas un chevauchement genant avec une forme
                    // voisine, ce qui est la seule raison d'etre du retrecissement force (voir
                    // Javadoc de la classe). On repart donc de la taille d'origine et on ne
                    // retrecit que jusqu'a la disparition de la collision reelle detectee plus
                    // haut - un debordement residuel au-dela de la boite, mais sans chevaucher
                    // la forme voisine, est accepte tel quel.
                    for (Map.Entry<XSLFTextRun, Double> e : baseline.entrySet()) {
                        e.getKey().setFontSize(e.getValue());
                    }
                    VerticalAlignment valign = ts.getVerticalAlignment();
                    factor = 1.0;
                    iter = 0;
                    didShrink = false;
                    textHeight = visibleMeasurer != null ? visibleMeasurer.height() : ts.getTextHeight(graphics);
                    List<Rectangle2D> zones = computeOverflowZones(anchor, textHeight, valign);
                    boolean stillColliding = !zones.isEmpty() && findCollidingShape(zones, ts, allTextShapes) != null;

                    while (stillColliding && factor > MIN_SCALE && iter < MAX_ITER) {
                        factor -= STEP;
                        for (Map.Entry<XSLFTextRun, Double> e : baseline.entrySet()) {
                            e.getKey().setFontSize(Math.max(1.0, e.getValue() * factor));
                        }
                        textHeight = visibleMeasurer != null ? visibleMeasurer.height() : ts.getTextHeight(graphics);
                        zones = computeOverflowZones(anchor, textHeight, valign);
                        stillColliding = !zones.isEmpty() && findCollidingShape(zones, ts, allTextShapes) != null;
                        iter++;
                        didShrink = true;
                    }

                    if (stillColliding) {
                        // Meme une reduction jusqu'a la limite basse ne suffit pas a ecarter la
                        // collision : la taille d'origine est restauree plutot que de produire
                        // un texte ecrase qui chevauche quand meme la forme voisine.
                        for (Map.Entry<XSLFTextRun, Double> e : baseline.entrySet()) {
                            e.getKey().setFontSize(e.getValue());
                        }
                        if (LOG.isDebugEnabled()) {
                            LOG.debug("{} : aucun ajustement (complet ou partiel jusqu'a disparition de la collision) "
                                    + "n'a ete possible -> boite structurellement trop petite pour ce texte, taille "
                                    + "d'origine restauree", shape.getShapeName());
                        }
                        continue;
                    }
                    if (didShrink && LOG.isDebugEnabled()) {
                        LOG.debug("{} : ajustement complet dans la boite impossible, retrecie a {}% seulement jusqu'a "
                                + "disparition de la collision (hauteur texte {}pt, boite {}pt) - debordement residuel "
                                + "au-dela de la boite accepte, comme le ferait PowerPoint pour ce mode noAutofit",
                                shape.getShapeName(), Math.round(factor * 100), textHeight, anchor.getHeight());
                    }
                }

                if (didShrink) {
                    count++;
                    if (LOG.isDebugEnabled()) {
                        String forcedSuffix = forcedByDeclaredNone
                                ? " [force : autofit=NONE dans le fichier d'origine]"
                                : "";
                        LOG.debug("{} retrecie a {}% (hauteur texte {}pt -> boite {}pt){}",
                                shape.getShapeName(), Math.round(factor * 100), textHeight, anchor.getHeight(), forcedSuffix);
                    }
                }
            } finally {
                if (visibleMeasurer != null) {
                    visibleMeasurer.close();
                }
            }
        }
        return count;
    }

    /**
     * Plus grande taille de police <em>declaree</em> (avant toute reduction)
     * parmi les runs non vides de la forme, ou -1 si aucune n'a de taille
     * explicite (heritee du theme/layout, non determinable ici sans parcourir
     * la chaine d'heritage complete).
     */
    private static double maxDeclaredFontSize(XSLFTextShape ts) {
        double max = -1;
        for (XSLFTextParagraph para : ts.getTextParagraphs()) {
            for (XSLFTextRun run : para.getTextRuns()) {
                String text = run.getRawText();
                if (text == null || text.isEmpty()) {
                    continue;
                }
                Double size = run.getFontSize();
                if (size != null && size > max) {
                    max = size;
                }
            }
        }
        return max;
    }

    /**
     * Visibilite package (pas {@code private}) : reutilisee telle quelle par {@link
     * TableCellLineSpacingFixer} pour capturer/restaurer les tailles de police d'une
     * cellule de tableau, meme raison que {@link #STEP}.
     */
    static Map<XSLFTextRun, Double> captureBaselineFontSizes(XSLFTextShape ts) {
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

    /** Un mot (sequence non blanche) - voir {@link #maxBoldWidthPremium}. */
    private static final Pattern WORD_PATTERN = Pattern.compile("\\S+");

    /**
     * {@code true} si, a la largeur utile ACTUELLE de {@code ts} reduite de
     * {@link #WIDTH_SAFETY_MARGIN} ET, le cas echeant, de {@link
     * #maxBoldWidthPremium}, au moins un mot bascule sur une nouvelle ligne
     * par rapport a la largeur pleine de {@code anchor} - signe qu'un point
     * de retour a la ligne se situe pres du bord droit de l'ancre (voir
     * Javadoc de la classe, section "Quatrieme garde-fou..."). Mesure via
     * l'ancre REELLE de {@code ts} (mutee temporairement puis aussitot
     * restauree) plutot que via une copie jetable : {@link
     * XSLFTextShape#getTextHeight(Graphics2D)} ignore deja la hauteur de
     * l'ancre pour son calcul, et comparer deux mesures sur le MEME objet
     * (plutot qu'objet source vs copie) evite par construction tout ecart de
     * modele de hauteur entre les deux mesures - voir "Amendement (2e) :
     * retour a la mesure sur le meme objet" en Javadoc de la classe pour
     * l'historique de cette decision.
     */
    static boolean isWrapMarginUnstable(XSLFTextShape ts, Rectangle2D anchor, Graphics2D graphics) {
        double fullWidthHeight = ts.getTextHeight(graphics);
        double reducedWidth = anchor.getWidth() * WIDTH_SAFETY_MARGIN - maxBoldWidthPremium(ts, graphics);
        if (reducedWidth < 1.0) {
            reducedWidth = 1.0;
        }
        Rectangle2D reduced = new Rectangle2D.Double(anchor.getX(), anchor.getY(), reducedWidth, anchor.getHeight());
        ts.setAnchor(reduced);
        double reducedWidthHeight = ts.getTextHeight(graphics);
        ts.setAnchor(anchor);
        return reducedWidthHeight > fullWidthHeight + 0.01;
    }

    /**
     * Plus grand ecart de largeur, parmi tous les mots des runs EN GRAS de
     * {@code ts}, entre la largeur reelle du mot (police en gras, celle
     * effectivement dessinee) et la largeur qu'aurait ce meme mot dans la
     * MEME police mais SANS le gras - 0 si {@code ts} ne contient aucun run
     * en gras. Voir Javadoc de la classe, section "Quatrieme garde-fou...",
     * amendement du 2026-09-02 : cible directement l'hypothese diagnostiquee
     * sur "ZoneTexte 41" (le calcul de decoupe de POI sous-estimant la
     * largeur du gras), en ajoutant cet ecart, mesure via {@code
     * Graphics2D.getFontMetrics(Font)}, a la marge de securite generique
     * {@link #WIDTH_SAFETY_MARGIN} lors de la reduction de largeur testee par
     * {@link #isWrapMarginUnstable}. Pour une forme sans aucun run en gras,
     * retourne 0 : le comportement redevient alors EXACTEMENT celui de la
     * marge generique seule (aucun risque de regression sur les formes non
     * concernees par cette hypothese).
     */
    /**
     * Corrige la fidelite visuelle de l'interligne pour les paragraphes declarant un
     * interligne en POURCENTAGE ({@code spcPct}, {@code getLineSpacing() > 0}) lorsque la
     * police effectivement utilisee ne declare aucun "leading" natif ({@code
     * java.awt.font.TextLayout.getLeading() == 0}, cas frequent - voir Javadoc de la classe,
     * "Cinquieme garde-fou...") : dans ce cas, Apache POI ne dessine pas les lignes a
     * l'interligne declare mais a un interligne gonfle par son propre repli interne de +15% de
     * (ascent+descent) - voir {@code org.apache.poi.sl.draw.DrawTextFragment#getLeading()} dans
     * le code source de POI 5.2.5, dont le commentaire indique explicitement que cette valeur a
     * ete choisie pour se rapprocher de LibreOffice/OpenOffice, pas de PowerPoint.
     *
     * <p>Plutot que de reimplementer tout le calcul de mise en page (retour a la ligne,
     * positionnement) hors de POI, cette methode reutilise le mecanisme de dessin natif de POI
     * en lui fournissant, pour les paragraphes concernes, une valeur d'interligne CORRIGEE
     * (recalculee pour compenser exactement le repli interne +15% de POI et retomber sur {@link
     * #LINE_HEIGHT_FIDELITY_FACTOR}) de sorte que le resultat VISUEL de {@code slide.draw()} se
     * rapproche de PowerPoint - et, en consequence, que {@code ts.getTextHeight(graphics)},
     * utilise par tout le reste de cette classe, redevienne lui aussi fidele sans calcul
     * separe.
     *
     * <p>Portee assumee (voir Javadoc de la classe pour le detail complet) : (1) une seule
     * valeur d'interligne s'applique par paragraphe ; quand plusieurs runs d'un meme paragraphe
     * declarent des tailles differentes, cette methode se cale sur le run de plus grande taille
     * (celui qui domine la hauteur de la ligne dans l'algorithme reel de POI) - les runs plus
     * petits du meme paragraphe restent une approximation. (2) la correction ne s'applique qu'
     * aux runs dont {@code TextLayout.getLeading()} resout a 0 pour le run retenu ; une police
     * declarant son propre leading natif n'est pas connue pour souffrir du meme ecart et n'est
     * pas touchee, faute de donnee reelle sur ce cas. (3) limite connue, non corrigee : le
     * modele de POI ne multiplie par le %% d'interligne que l'AVANCE ENTRE les lignes, jamais la
     * hauteur propre de la toute premiere ligne du bloc (toujours son ascent+descent+leading
     * brut, non compense) - l'erreur residuelle sur cette seule ligne reste marginale sur un
     * bloc de plusieurs lignes, mais cette methode ne la corrige pas.
     */
    static void correctPercentLineSpacingForFidelity(XSLFTextShape ts, Graphics2D graphics) {
        FontRenderContext frc = graphics.getFontRenderContext();
        for (XSLFTextParagraph para : ts.getTextParagraphs()) {
            Double declaredSpacing = para.getLineSpacing();
            if (declaredSpacing == null || declaredSpacing <= 0) {
                continue; // interligne absolu (points) ou non defini localement : hors de portee
            }
            XSLFTextRun governingRun = null;
            double governingSize = -1;
            for (XSLFTextRun run : para.getTextRuns()) {
                String text = run.getRawText();
                if (text == null || text.trim().isEmpty()) {
                    continue;
                }
                Double size = run.getFontSize();
                if (size != null && size > governingSize) {
                    governingSize = size;
                    governingRun = run;
                }
            }
            if (governingRun == null) {
                continue; // paragraphe vide ou tailles non resolues localement : hors de portee
            }
            String family = governingRun.getFontFamily();
            String resolvedFamily = family != null ? family : DEFAULT_FONT_FAMILY;
            int style = (governingRun.isBold() ? Font.BOLD : 0) | (governingRun.isItalic() ? Font.ITALIC : 0);
            Font font = new Font(resolvedFamily, style, 12).deriveFont((float) governingSize);
            TextLayout layout = new TextLayout(governingRun.getRawText(), font, frc);
            if (layout.getLeading() != 0) {
                continue; // la police declare son propre leading : le repli +15% incrimine de POI ne s'applique pas ici
            }
            double ascentPlusDescent = layout.getAscent() + layout.getDescent();
            if (ascentPlusDescent <= 0) {
                continue;
            }
            double poiLineHeight = ascentPlusDescent * 1.15; // repli de POI, voir DrawTextFragment.getLeading()
            double targetLineHeight = governingSize * LINE_HEIGHT_FIDELITY_FACTOR;
            double correctedSpacing = declaredSpacing * (targetLineHeight / poiLineHeight);
            para.setLineSpacing(correctedSpacing);
            if (LOG.isDebugEnabled()) {
                LOG.debug("{} : interligne corrige de {}% a {}% pour la fidelite visuelle (police '{}' {}pt, "
                        + "ascent+descent={}, repli POI suppose {}, cible {}, voir Javadoc de la classe)",
                        ts.getShapeName(), declaredSpacing, correctedSpacing, font.getFamily(), governingSize,
                        ascentPlusDescent, poiLineHeight, targetLineHeight);
            }
        }
    }

    private static double maxBoldWidthPremium(XSLFTextShape ts, Graphics2D graphics) {
        double maxPremium = 0;
        for (XSLFTextParagraph para : ts.getTextParagraphs()) {
            for (XSLFTextRun run : para.getTextRuns()) {
                if (!run.isBold()) {
                    continue;
                }
                String text = run.getRawText();
                if (text == null || text.isEmpty()) {
                    continue;
                }
                Double declaredSize = run.getFontSize();
                if (declaredSize == null) {
                    continue; // taille non resolue localement ici : hors de portee, voir Javadoc de la classe
                }
                String family = run.getFontFamily();
                String resolvedFamily = family != null ? family : DEFAULT_FONT_FAMILY;
                int styleWithBold = Font.BOLD | (run.isItalic() ? Font.ITALIC : Font.PLAIN);
                int styleWithoutBold = run.isItalic() ? Font.ITALIC : Font.PLAIN;
                Font boldFont = new Font(resolvedFamily, styleWithBold, 12).deriveFont(declaredSize.floatValue());
                Font plainFont = new Font(resolvedFamily, styleWithoutBold, 12).deriveFont(declaredSize.floatValue());
                FontMetrics boldMetrics = graphics.getFontMetrics(boldFont);
                FontMetrics plainMetrics = graphics.getFontMetrics(plainFont);

                Matcher m = WORD_PATTERN.matcher(text);
                while (m.find()) {
                    String word = m.group();
                    double premium = boldMetrics.stringWidth(word) - plainMetrics.stringWidth(word);
                    if (premium > maxPremium) {
                        maxPremium = premium;
                    }
                }
            }
        }
        return maxPremium;
    }

    /**
     * Calcule la ou les zones de "debordement" - au-dela de l'anchor
     * d'origine d'une forme - qu'occuperait son texte s'il n'etait pas
     * retreci, selon son alignement vertical (TOP/MIDDLE/BOTTOM ; TOP par
     * defaut si non precise, convention OOXML {@code anchor="t"}). Retourne
     * une liste vide si le texte ne deborde pas.
     */
    static List<Rectangle2D> computeOverflowZones(Rectangle2D anchor, double textHeight, VerticalAlignment valign) {
        List<Rectangle2D> zones = new ArrayList<>();
        double excess = textHeight - anchor.getHeight();
        if (excess <= 0) {
            return zones;
        }
        VerticalAlignment v = (valign == null) ? VerticalAlignment.TOP : valign;
        switch (v) {
            case BOTTOM:
                zones.add(new Rectangle2D.Double(anchor.getX(), anchor.getY() - excess, anchor.getWidth(), excess));
                break;
            case MIDDLE:
                double half = excess / 2.0;
                zones.add(new Rectangle2D.Double(anchor.getX(), anchor.getY() - half, anchor.getWidth(), half));
                zones.add(new Rectangle2D.Double(anchor.getX(), anchor.getY() + anchor.getHeight(), anchor.getWidth(), half));
                break;
            case TOP:
            default:
                zones.add(new Rectangle2D.Double(anchor.getX(), anchor.getY() + anchor.getHeight(), anchor.getWidth(), excess));
                break;
        }
        return zones;
    }

    /**
     * Indique si l'une des zones de debordement fournies chevauche l'anchor
     * d'une autre forme de texte non vide du slide. Les formes sans texte
     * (rectangles/panneaux de fond) sont ignorees : un debordement dessus
     * n'occasionne aucune confusion visuelle.
     */
    static boolean overflowCollidesWithText(List<Rectangle2D> overflowZones, XSLFTextShape self, List<XSLFShape> allTextShapes) {
        return findCollidingShape(overflowZones, self, allTextShapes) != null;
    }

    /**
     * {@code true} si {@code shape} est une forme de type "bulle" (callout
     * OOXML - ex. "Bulle narrative : rectangle a coins arrondis", {@link
     * ShapeType#WEDGE_ROUND_RECT_CALLOUT}) : couvre toute la famille de types
     * dont le nom contient {@code "CALLOUT"} dans {@link ShapeType} ({@code
     * CALLOUT_1..3}, variantes "accent"/"border", fleches de rappel, {@code
     * WEDGE_*_CALLOUT}, {@code CLOUD_CALLOUT}, etc.). Voir Javadoc de la
     * classe, section "Exemption des bulles narratives (callouts) de tout
     * chevauchement" pour la justification complete. Partagee
     * (package-private) avec {@link NeighborShapeOverlapFixer}.
     */
    static boolean isCalloutShape(XSLFShape shape) {
        if (!(shape instanceof XSLFTextShape)) {
            return false;
        }
        ShapeType type = ((XSLFTextShape) shape).getShapeType();
        return type != null && type.name().contains("CALLOUT");
    }

    /**
     * Comme {@link #overflowCollidesWithText}, mais retourne la forme
     * responsable de la collision (ou {@code null}) plutot qu'un simple
     * booleen - utilise pour enrichir les logs de diagnostic avec le nom de
     * la forme en cause.
     */
    private static XSLFShape findCollidingShape(List<Rectangle2D> overflowZones, XSLFTextShape self, List<XSLFShape> allTextShapes) {
        for (Rectangle2D zone : overflowZones) {
            for (XSLFShape other : allTextShapes) {
                if (other == self || !(other instanceof XSLFTextShape) || isCalloutShape(other)) {
                    continue;
                }
                XSLFTextShape ots = (XSLFTextShape) other;
                String text = ots.getText();
                if (text == null || text.trim().isEmpty()) {
                    continue;
                }
                Rectangle2D otherAnchor = ots.getAnchor();
                if (otherAnchor != null && zone.intersects(otherAnchor)) {
                    return other;
                }
            }
        }
        return null;
    }

    /** Parcourt les formes du slide, y compris a l'interieur des groupes, et ne garde que celles porteuses de texte. */
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

    /**
     * {@code true} si la forme contient au moins un paragraphe VISIBLE ET si
     * son DERNIER paragraphe est entierement VIDE - le seul cas ou {@link
     * VisibleTextMeasurer} change quoi que ce soit a la mesure (voir
     * "Troisieme garde-fou" et son amendement du 2026-09-01 en Javadoc de
     * classe). Une forme entierement vide (aucun paragraphe visible) est
     * exclue par construction : rien ne la distingue alors du comportement
     * d'avant ce correctif, {@code ts.getTextHeight()} reste utilisee
     * directement.
     *
     * <p>Deliberement plus etroit que "contient un paragraphe vide quelque
     * part" (comportement d'avant l'amendement du 2026-09-01) : un paragraphe
     * vide EN MILIEU de bloc (au moins un paragraphe visible le suit) occupe
     * un espace reel qui repousse ce qui le suit dans le rendu reel - il ne
     * doit jamais etre ignore de la mesure, voir Javadoc de la classe. Seule
     * une sequence finale de paragraphes vides (le dernier paragraphe du bloc,
     * et potentiellement ceux juste avant lui s'ils sont eux aussi vides) ne
     * "pousse" plus rien apres elle et peut donc etre ignoree sans
     * consequence sur le positionnement du texte visible reellement dessine.
     */
    private static boolean hasTrailingBlankParagraph(XSLFTextShape ts) {
        List<XSLFTextParagraph> paragraphs = ts.getTextParagraphs();
        if (paragraphs.isEmpty() || isVisibleParagraph(paragraphs.get(paragraphs.size() - 1))) {
            return false;
        }
        for (XSLFTextParagraph para : paragraphs) {
            if (isVisibleParagraph(para)) {
                return true; // au moins un paragraphe visible existe ailleurs dans le bloc
            }
        }
        return false; // forme entierement vide : hors de portee de ce garde-fou
    }

    /**
     * {@code true} si {@code ts} declare EXPLICITEMENT au moins un des trois
     * modes d'autofit ({@code noAutofit}/{@code normAutofit}/{@code
     * spAutoFit}) dans son PROPRE {@code <a:bodyPr>} au niveau de la slide -
     * {@code false} si rien n'y est declare (bodyPr absent, ou present mais
     * vide), ce qui correspond au motif OOXML normal ou l'autofit est
     * simplement herite de la mise en page/du masque sans etre redeclare
     * localement.
     *
     * <p>Contrairement a {@link XSLFTextShape#getTextAutofit()}, qui ne
     * distingue pas ce cas d'un {@code normAutofit} explicitement voulu et
     * retombe dans les deux cas sur {@code NORMAL} (voir Javadoc de la
     * classe, section "Elargissement cible..."), lit directement le XML brut
     * via {@link XSLFShape#getXmlObject()} (public) pour faire cette
     * distinction - {@code getTextBodyPr()} de POI, qui ferait la meme
     * distinction via {@code isSetNoAutofit()}/{@code isSetNormAutofit()}/
     * {@code isSetSpAutoFit()}, n'est que {@code protected} et donc
     * inaccessible depuis ce paquetage.
     */
    static boolean isAutofitExplicitlyDeclaredLocally(XSLFTextShape ts) {
        Object raw = ts.getXmlObject();
        if (!(raw instanceof CTShape)) {
            // Forme non standard (rare - normalement toujours <p:sp> pour une
            // XSLFTextShape) : indeterminable, par prudence l'autofit est considere
            // comme declare (comportement inchange, pas d'elargissement).
            return true;
        }
        CTTextBody txBody = ((CTShape) raw).getTxBody();
        CTTextBodyProperties bodyPr = txBody == null ? null : txBody.getBodyPr();
        if (bodyPr == null) {
            return false;
        }
        return bodyPr.isSetNoAutofit() || bodyPr.isSetNormAutofit() || bodyPr.isSetSpAutoFit();
    }

    /**
     * {@code true} si {@code slide} est detectee comme une diapositive de
     * sommaire/table des matieres (voir Javadoc de la classe, section
     * "Elargissement cible...") : sa forme de titre (voir {@link
     * TitleDetector#findTitleShape}) existe et son texte correspond a un
     * intitule connu (voir {@link #isTableOfContentsTitle}).
     */
    static boolean isSommaireSlide(XSLFSlide slide) {
        XSLFTextShape titleShape = TitleDetector.findTitleShape(slide);
        if (titleShape == null) {
            return false;
        }
        String title = titleShape.getText();
        boolean matches = isTableOfContentsTitle(title);
        if (matches && LOG.isDebugEnabled()) {
            LOG.debug("Slide detectee comme sommaire/table des matieres (titre : '{}') -> elargissement "
                    + "cible du retrecissement force (voir Javadoc de la classe)", title);
        }
        return matches;
    }

    /**
     * {@code true} si {@code ts}, sur une diapositive pour laquelle {@code
     * sommaireSlide} (calcule une seule fois par slide via {@link
     * #isSommaireSlide}, voir {@link #fitOverflowingText}) vaut {@code true},
     * releve de l'elargissement cible aux diapositives sommaire/table des
     * matieres (voir Javadoc de la classe, section "Elargissement cible...") :
     * son autofit n'est NI {@code NONE} explicitement declare (auquel cas
     * {@link #fitOverflowingText} la traite deja via {@code
     * forcedByDeclaredNone}, sans rapport avec cet elargissement) NI
     * explicitement declare localement d'une autre maniere (voir {@link
     * #isAutofitExplicitlyDeclaredLocally}).
     *
     * <p>Partagee avec {@link NeighborShapeOverlapFixer} (package-private,
     * meme paquetage) : {@link #fitOverflowingText} ne retrecit JAMAIS une
     * telle forme (voir la sortie anticipee dans sa boucle), y compris quand
     * son texte depasse sa propre ancre - a la difference de toute autre forme
     * de ce cas (traitee, elle, par {@link #fitOverflowingText}), qui reste
     * donc hors de portee de {@link NeighborShapeOverlapFixer} tant qu'elle
     * depasse sa propre ancre (voir la Javadoc de cette classe, section
     * "Complementaire, jamais redondant"). Une forme exemptee par cette
     * methode a besoin d'un traitement (reduction d'interligne notamment) si
     * elle chevauche une forme voisine meme en depassant sa propre ancre -
     * {@link NeighborShapeOverlapFixer} l'identifie via cette meme methode
     * pour lever, dans ce seul cas, son propre garde-fou "deja gere par
     * OverflowAwareTextFitter".
     */
    static boolean isSommaireBroadeningExempt(XSLFTextShape ts, boolean sommaireSlide) {
        return sommaireSlide
                && ts.getTextAutofit() != TextShape.TextAutofit.NONE
                && !isAutofitExplicitlyDeclaredLocally(ts);
    }

    /**
     * {@code true} si AUCUN maillon de la chaine d'heritage complete de {@code
     * ts} - la forme sur la slide, PUIS son placeholder correspondant sur la
     * mise en page, PUIS sur le masque (voir {@code XSLFShape#fetchShapeProperty},
     * meme mecanisme public que {@link TitlePlaceholderResolver#resolve}) - ne
     * declare EXPLICITEMENT l'un des trois modes d'autofit dans son PROPRE
     * {@code <a:bodyPr>} (voir {@link #isAutofitExplicitlyDeclaredLocally},
     * applique ici a CHAQUE maillon, pas seulement a la forme de depart).
     *
     * <p>Contrairement a {@link #isAutofitExplicitlyDeclaredLocally} seule (qui
     * ne regarde que la forme sur la slide, et peut donc cacher une intention
     * d'autofit exprimee plus haut dans la chaine), un resultat {@code true}
     * ici signifie que l'autofit est reellement absent PARTOUT - la conclusion
     * "PowerPoint n'applique aucun retrecissement automatique pour cette forme"
     * est alors fiable, independamment du titre de la diapositive. Voir la
     * Javadoc de la classe, section "Elargissement general (experimental)",
     * pour le cas reel ayant motive cette methode et son statut experimental
     * (controle par {@link io.github.atlan77c.pptx2picture.RenderOptions
     * RenderOptions}{@code .isBroadenAutofitExemption()}).
     */
    static boolean isAutofitUndeclaredThroughoutInheritanceChain(XSLFTextShape ts) {
        PropertyFetcher<Boolean> fetcher = new PropertyFetcher<Boolean>() {
            @Override
            public boolean fetch(XSLFShape candidate) {
                if (!(candidate instanceof XSLFTextShape)) {
                    return false;
                }
                if (isAutofitExplicitlyDeclaredLocally((XSLFTextShape) candidate)) {
                    setValue(Boolean.TRUE);
                    return true;
                }
                // Rien declare a ce maillon (slide, mise en page ou masque) : on
                // continue vers le maillon suivant plutot que de conclure ici.
                return false;
            }
        };
        ts.fetchShapeProperty(fetcher);
        return !fetcher.isSet();
    }

    /**
     * Exemption combinee de retrecissement force pour cause d'autofit mal
     * classe par POI - vraie si {@code ts} releve soit de l'elargissement
     * "sommaire" existant ({@link #isSommaireBroadeningExempt}), soit, quand
     * {@code broadenAutofitExemption} est actif (voir {@link
     * io.github.atlan77c.pptx2picture.RenderOptions.Builder#broadenAutofitExemption(boolean)}),
     * du nouvel elargissement general independant du titre (voir {@link
     * #isAutofitUndeclaredThroughoutInheritanceChain} et la Javadoc de la
     * classe, section "Elargissement general (experimental)"). Les deux
     * exemptions restent actives en parallele : desactiver
     * {@code broadenAutofitExemption} restaure exactement le comportement
     * anterieur a cet ajout (2026-09-05).
     */
    static boolean isAutofitBroadeningExempt(XSLFTextShape ts, boolean sommaireSlide, boolean broadenAutofitExemption) {
        if (isSommaireBroadeningExempt(ts, sommaireSlide)) {
            return true;
        }
        return broadenAutofitExemption
                && ts.getTextAutofit() != TextShape.TextAutofit.NONE
                && isAutofitUndeclaredThroughoutInheritanceChain(ts);
    }

    /** {@code true} si {@code title}, une fois normalise (voir {@link #normalizeTitle}), figure parmi {@link #TABLE_OF_CONTENTS_TITLES}. */
    private static boolean isTableOfContentsTitle(String title) {
        return TABLE_OF_CONTENTS_TITLES.contains(normalizeTitle(title));
    }

    /**
     * Normalise un texte de titre pour comparaison : espaces de bordure
     * retires, accents/diacritiques retires (decomposition Unicode NFD puis
     * suppression des marques combinantes), passage en minuscules (locale
     * neutre - ces intitules sont des mots simples, sans regle de casse
     * dependante de la langue qui poserait probleme ici).
     */
    private static String normalizeTitle(String text) {
        if (text == null) {
            return "";
        }
        String decomposed = Normalizer.normalize(text.trim(), Normalizer.Form.NFD);
        String withoutAccents = decomposed.replaceAll("\\p{M}", "");
        return withoutAccents.toLowerCase(Locale.ROOT);
    }

    private static Set<String> buildNormalizedTitleSet(String... titles) {
        Set<String> set = new HashSet<>();
        for (String title : titles) {
            set.add(normalizeTitle(title));
        }
        return set;
    }

    /**
     * Copie de mesure jetable omettant uniquement la SEQUENCE FINALE de
     * paragraphes entierement vides d'une forme (voir {@link
     * #hasTrailingBlankParagraph}, seul cas ou cette classe est utilisee) -
     * ces paragraphes ne portent aucun contenu visible mais {@code
     * ts.getTextHeight(graphics)} les compte comme des lignes de texte
     * reelles (voir "Troisieme garde-fou" en Javadoc de classe). Un paragraphe
     * vide EN MILIEU de bloc (au moins un paragraphe visible le suit) reste,
     * lui, recopie normalement avec sa hauteur reelle : contrairement a un
     * paragraphe vide en fin de bloc, il occupe un espace reel qui repousse
     * vers le bas le texte visible qui le suit dans le rendu reel de POI - l'
     * omettre sous-estimerait la hauteur reellement necessaire (voir
     * l'amendement du 2026-09-01 en Javadoc de classe, motive par le fichier
     * "fichier-test-A.pptx"). Meme technique que {@code
     * NeighborShapeOverlapFixer.ParagraphMeasurer} (copie {@link XSLFTextBox}
     * jetable, jamais dessinee, ajoutee au meme slide puis retiree par {@link
     * #close()}), simplifiee ici a un seul total a mesurer (pas de
     * repartition par paragraphe) : la copie est construite une seule fois
     * par forme analysee, {@link #height()} resynchronisant ensuite les
     * tailles de police - seule propriete que {@link #fitOverflowingText}
     * fait varier au fil des iterations - sur les valeurs COURANTES des runs
     * reels avant de mesurer.
     */
    private static final class VisibleTextMeasurer implements AutoCloseable {

        private final XSLFSlide slide;
        private final XSLFTextBox scratch;
        private final Graphics2D graphics;
        private final List<XSLFTextRun> sourceRuns = new ArrayList<>();
        private final List<XSLFTextRun> scratchRuns = new ArrayList<>();

        VisibleTextMeasurer(XSLFSlide slide, XSLFTextShape source, Rectangle2D anchor, Graphics2D graphics) {
            this.slide = slide;
            this.graphics = graphics;

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
            // suite, seuls les paragraphes retenus par "keep" ci-dessous sont recopies.
            scratch.removeTextParagraph(scratch.getTextParagraphs().get(0));

            // Determine, pour chaque paragraphe source, s'il doit etre recopie : tout
            // paragraphe VISIBLE, plus tout paragraphe VIDE suivi (plus loin dans le bloc)
            // d'au moins un paragraphe visible (paragraphe vide EN MILIEU de bloc - conserve
            // sa contribution reelle a la position de ce qui le suit, voir Javadoc de la
            // classe). Seule la sequence finale de paragraphes vides (aucun paragraphe
            // visible apres eux) est exclue.
            List<XSLFTextParagraph> paragraphs = source.getTextParagraphs();
            boolean[] keep = new boolean[paragraphs.size()];
            boolean visibleFollows = false;
            for (int i = paragraphs.size() - 1; i >= 0; i--) {
                boolean visible = isVisibleParagraph(paragraphs.get(i));
                keep[i] = visible || visibleFollows;
                if (visible) {
                    visibleFollows = true;
                }
            }

            for (int i = 0; i < paragraphs.size(); i++) {
                if (!keep[i]) {
                    continue; // paragraphe vide de la sequence finale : deliberement exclu de la copie
                }
                XSLFTextParagraph para = paragraphs.get(i);
                XSLFTextParagraph copy = scratch.addNewTextParagraph();
                // Amendement (2026-09-02) : getLineSpacing()/getSpaceBefore()/getSpaceAfter()
                // renvoient un Double NULLABLE (non defini localement sur ce paragraphe
                // source -> herite du layout/masque), alors que les setters correspondants
                // prennent un double PRIMITIF - un appel direct ici, non garde, deballait
                // (unboxing) un null et levait une NullPointerException des qu'un paragraphe
                // ne declarait pas explicitement son propre espacement (motif tres courant,
                // decouvert via un echec de mvn verify reel sur "fichier-test-A.pptx" apres
                // l'introduction, ailleurs dans cette classe, d'un nouveau code (depuis
                // retire, voir Javadoc de la classe, section "Quatrieme garde-fou...")
                // copiant ce meme motif non protege. Ne fixer la
                // propriete que si une valeur EXPLICITE existe ; sinon laisser le paragraphe
                // de la copie a son propre defaut (comportement de mesure inchange : un
                // paragraphe neuf sans valeur fixee retombe deja sur le meme defaut que POI
                // utiliserait pour le paragraphe source qui n'en declare pas non plus).
                Double paraLineSpacing = para.getLineSpacing();
                if (paraLineSpacing != null) {
                    copy.setLineSpacing(paraLineSpacing);
                }
                Double paraSpaceBefore = para.getSpaceBefore();
                if (paraSpaceBefore != null) {
                    copy.setSpaceBefore(paraSpaceBefore);
                }
                Double paraSpaceAfter = para.getSpaceAfter();
                if (paraSpaceAfter != null) {
                    copy.setSpaceAfter(paraSpaceAfter);
                }
                for (XSLFTextRun run : para.getTextRuns()) {
                    XSLFTextRun runCopy = copy.addNewTextRun();
                    String text = run.getRawText();
                    runCopy.setText(text == null ? "" : text);
                    Double size = run.getFontSize();
                    if (size != null) {
                        runCopy.setFontSize(size);
                    }
                    String family = run.getFontFamily();
                    if (family != null) {
                        runCopy.setFontFamily(family);
                    }
                    sourceRuns.add(run);
                    scratchRuns.add(runCopy);
                }
            }
        }

        /** Hauteur mesuree du texte visible, apres resynchronisation des tailles de police courantes des runs reels. */
        double height() {
            for (int i = 0; i < sourceRuns.size(); i++) {
                Double size = sourceRuns.get(i).getFontSize();
                if (size != null) {
                    scratchRuns.get(i).setFontSize(size);
                }
            }
            return scratch.getTextHeight(graphics);
        }

        @Override
        public void close() {
            slide.removeShape(scratch);
        }
    }
}
