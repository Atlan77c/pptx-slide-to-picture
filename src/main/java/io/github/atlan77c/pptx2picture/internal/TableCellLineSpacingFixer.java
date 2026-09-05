package io.github.atlan77c.pptx2picture.internal;

import org.apache.poi.xslf.usermodel.XSLFGroupShape;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTable;
import org.apache.poi.xslf.usermodel.XSLFTableCell;
import org.apache.poi.xslf.usermodel.XSLFTableRow;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.awt.Graphics2D;
import java.awt.geom.Rectangle2D;
import java.util.ArrayList;
import java.util.Collections;
import java.util.IdentityHashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;

/**
 * Applique aux cellules d'un tableau natif ({@link XSLFTable}) le meme correctif de
 * fidelite d'interligne que {@link OverflowAwareTextFitter#correctPercentLineSpacingForFidelity}
 * applique deja aux formes de texte autonomes - voir la section "Hors de portee : les
 * cellules de tableau natif" de la Javadoc de cette derniere pour le constat complet.
 *
 * <h2>Cas reel a l'origine de ce correctif (2026-09-04)</h2>
 * <p>Signale par l'utilisateur : fichier "Mes Evenements Emploi et Prestation - Doc
 * vision 1.0.pptx", slide 18, tableau "Tableau 6". Une cellule de la colonne "Analyse"
 * (3 paragraphes, aucun {@code <a:pPr>} local - interligne entierement herite, donc en
 * pourcentage) affichait un texte visiblement plus haut que chez PowerPoint, au point de
 * chevaucher la cellule de la ligne suivante dans la meme colonne. Diagnostic de
 * l'utilisateur, confirme par lecture du XML et du code source POI 5.2.5 (miroir GitHub
 * officiel) : pas un veritable chevauchement de mise en page, mais le meme repli interne
 * +15% de POI (au lieu du modele reel de PowerPoint, {@code taille x
 * LINE_HEIGHT_FIDELITY_FACTOR}) deja documente et corrige pour les formes de texte
 * normales - jamais applique ici car aucun des correctifs existants ne parcourt les
 * cellules de tableau (voir Javadoc de {@link OverflowAwareTextFitter}, section
 * ci-dessus).
 *
 * <h2>Pourquoi une classe dediee plutot qu'etendre {@link OverflowAwareTextFitter}</h2>
 * <p>{@code XSLFTableCell} EST une {@link org.apache.poi.xslf.usermodel.XSLFTextShape}
 * (elle en herite directement), donc {@link
 * OverflowAwareTextFitter#correctPercentLineSpacingForFidelity} s'applique telle quelle,
 * sans aucune modification, a une cellule. Mais integrer les cellules au parcours
 * generique de {@link OverflowAwareTextFitter#fitOverflowingText} exposerait aussi les
 * cellules a ses AUTRES garde-fous (marge de retour a la ligne horizontale, exemption
 * sommaire/callout, mesure des paragraphes vides en fin de bloc via {@code
 * VisibleTextMeasurer}...), tous calibres et testes uniquement pour des formes autonomes
 * - jamais pour la geometrie particuliere d'un tableau (lignes qui grandissent
 * dynamiquement selon le contenu, cellules fusionnees sur plusieurs lignes/colonnes). Une
 * classe dediee, au perimetre volontairement plus etroit (uniquement le correctif
 * d'interligne, plus un retrecissement de police en toute derniere extremite), evite ce
 * risque de regression non maitrise sur l'ensemble des fichiers deja confirmes par
 * l'utilisateur.
 *
 * <h2>Pourquoi deux passes distinctes sur le meme tableau</h2>
 * <p>{@code XSLFTableCell#getAnchor()} declenche, a son PREMIER appel sur n'importe
 * quelle cellule du tableau, un recalcul de la geometrie de TOUTES ses cellules ({@code
 * XSLFTable#updateCellAnchor()}, qui grandit chaque ligne selon {@code
 * DrawTextShape#getTextHeight()} de ses cellules - voir Javadoc de {@link
 * OverflowAwareTextFitter}). Corriger l'interligne d'une cellule APRES avoir deja lu
 * l'ancre d'une autre cellule du meme tableau appellerait ce recalcul une seconde fois de
 * toute facon (voir plus bas), mais avec un risque d'ordre subtil si un appelant futur
 * lisait une ancre entre-temps pour un autre besoin. Cette classe fixe l'ordre
 * explicitement : {@link #fixTableCellLineSpacing} corrige d'abord l'interligne de
 * TOUTES les cellules du tableau (phase A), puis force explicitement {@code
 * table.updateCellAnchor()} une seule fois (recalcul a partir des interlignes deja
 * corriges), avant de ne mesurer/retrecir que si necessaire (phase B) - jamais l'inverse.
 *
 * <h2>Retrecissement de secours (phase B)</h2>
 * <p>Le correctif d'interligne seul (phase A) resout la tres grande majorite du
 * chevauchement mesure (c'est le mecanisme diagnostique reel, voir ci-dessus). Un residu
 * peut neanmoins subsister : {@code correctPercentLineSpacingForFidelity} ne corrige pas
 * la hauteur propre de la toute premiere ligne d'un bloc (limite deja documentee), et la
 * mesure utilisee par {@code updateCellAnchor()} pour agrandir la ligne ({@code
 * DrawTextShape#getTextHeight()}, sans {@link Graphics2D} explicite) peut differer
 * legerement du {@link Graphics2D} reel utilise pour le rendu final. Phase B mesure donc,
 * pour chaque cellule NON fusionnee ({@code getGridSpan()==1 && getRowSpan()==1} - comme
 * la premiere passe de {@code updateCellAnchor()} elle-meme, qui ignore les cellules
 * fusionnees pour le calcul de hauteur de ligne), si {@code cell.getTextHeight(graphics)}
 * depasse encore sa propre ancre (deja recalculee en phase A) et, si oui, retrecit sa
 * police par paliers - reutilisant {@link OverflowAwareTextFitter#STEP}, {@link
 * OverflowAwareTextFitter#MIN_SCALE}, {@link OverflowAwareTextFitter#MAX_ITER} et {@link
 * OverflowAwareTextFitter#SAFETY_MARGIN} (partages, package-private) plutot que dupliquer
 * ces valeurs calibrees - jusqu'a tenir dans {@code anchor.getHeight() * SAFETY_MARGIN},
 * sans notion de collision avec une forme voisine (a la difference du retrecissement
 * "force" de {@link OverflowAwareTextFitter} pour l'autofit {@code NONE}) : une cellule de
 * tableau n'a pas d'equivalent de debordement volontaire - PowerPoint agrandirait
 * simplement la ligne, ce que notre rendu ne fait pas dynamiquement au dessin (voir
 * Javadoc de {@link OverflowAwareTextFitter}, meme section) - donc tout depassement
 * residuel de sa propre ancre reste toujours un artefact de mesure a corriger.
 *
 * <p><b>Cellules fusionnees, portee non couverte</b> : ni la phase A (qui s'applique a
 * toutes les cellules, fusionnees ou non - la correction d'interligne reste toujours
 * benefique) ni la phase B (volontairement restreinte aux cellules simples, comme {@code
 * updateCellAnchor()} lui-meme) ne traitent le cas d'une cellule fusionnee dont le texte
 * deborderait encore apres correction d'interligne. Non rencontre a ce jour sur les
 * fichiers reels traites par ce projet.
 */
public final class TableCellLineSpacingFixer {

    private static final Logger LOG = LoggerFactory.getLogger(TableCellLineSpacingFixer.class);

    private TableCellLineSpacingFixer() {
    }

    /**
     * Corrige, in-place, l'interligne (et, en dernier recours, la taille de police) des
     * cellules de tableau du slide dont le texte visible deborderait sinon de sa propre
     * ligne - voir Javadoc de la classe. A appeler avant {@code slide.draw(graphics)}.
     *
     * @return le nombre de cellules dont la police a ete retrecie en phase B (voir Javadoc
     *         de la classe) - ne comptabilise PAS les corrections d'interligne de la phase
     *         A (qui n'a pas d'effet visuel mesurable independamment, seulement une
     *         hauteur mesuree corrigee), meme convention que {@link
     *         OverflowAwareTextFitter#fitOverflowingText}, dont le retour ne comptabilise
     *         pas non plus son propre cinquieme garde-fou (interligne).
     */
    public static int fixTableCellLineSpacing(XSLFSlide slide, Graphics2D graphics) {
        int count = 0;
        for (XSLFTable table : collectTables(slide.getShapes())) {
            count += fixTable(table, graphics);
        }
        return count;
    }

    private static int fixTable(XSLFTable table, Graphics2D graphics) {
        List<XSLFTableCell> cells = collectCellsWithVisibleText(table);
        if (cells.isEmpty()) {
            return 0;
        }

        // Phase A : corrige l'interligne de TOUTES les cellules du tableau avant de lire
        // la moindre ancre - voir Javadoc de la classe pour la raison de cet ordre.
        for (XSLFTableCell cell : cells) {
            OverflowAwareTextFitter.correctPercentLineSpacingForFidelity(cell, graphics);
        }

        // Force le recalcul de la geometrie du tableau a partir des interlignes deja
        // corriges (plutot que de compter sur le cache paresseux de getAnchor(), qui
        // pourrait avoir ete rempli plus tot par un autre correctif avec des valeurs
        // non corrigees) - voir Javadoc de la classe.
        table.updateCellAnchor();

        // Phase B : retrecissement de secours, uniquement pour les cellules simples
        // (voir Javadoc de la classe) dont le texte deborde encore de sa propre ligne
        // apres la correction d'interligne ci-dessus.
        int count = 0;
        for (XSLFTableCell cell : cells) {
            if (cell.getGridSpan() != 1 || cell.getRowSpan() != 1) {
                continue;
            }
            if (shrinkIfStillOverflowing(cell, graphics)) {
                count++;
            }
        }
        return count;
    }

    /**
     * {@code true} si la police de {@code cell} a ete retrecie pour tenir dans sa propre
     * ancre (deja recalculee par {@link XSLFTable#updateCellAnchor()} en phase A).
     * Meme technique que le retrecissement "systematique" (NORMAL/SHAPE) de {@link
     * OverflowAwareTextFitter#fitOverflowingText} - sans etape de detection de collision,
     * voir Javadoc de la classe pour la justification. Visibilite package (pas {@code
     * private}) : teste directement, sur une ancre de cellule fixee a la main, plutot que
     * via la geometrie dynamique complete d'un tableau (voir {@code
     * TableCellLineSpacingFixerTest}, meme convention que {@code
     * OverflowAwareTextFitter#correctPercentLineSpacingForFidelity}).
     */
    static boolean shrinkIfStillOverflowing(XSLFTableCell cell, Graphics2D graphics) {
        Rectangle2D anchor = cell.getAnchor();
        if (anchor == null || anchor.getHeight() <= 0) {
            return false;
        }

        double textHeight = cell.getTextHeight(graphics);
        double targetHeight = anchor.getHeight() * OverflowAwareTextFitter.SAFETY_MARGIN;
        if (textHeight <= targetHeight) {
            return false;
        }

        Map<XSLFTextRun, Double> baseline = OverflowAwareTextFitter.captureBaselineFontSizes(cell);
        if (baseline.isEmpty()) {
            return false; // tailles heritees du theme/layout, non modifiables ici
        }

        double factor = 1.0;
        int iter = 0;
        boolean didShrink = false;
        while (textHeight > targetHeight && factor > OverflowAwareTextFitter.MIN_SCALE
                && iter < OverflowAwareTextFitter.MAX_ITER) {
            factor -= OverflowAwareTextFitter.STEP;
            for (Map.Entry<XSLFTextRun, Double> e : baseline.entrySet()) {
                e.getKey().setFontSize(Math.max(1.0, e.getValue() * factor));
            }
            textHeight = cell.getTextHeight(graphics);
            iter++;
            didShrink = true;
        }

        if (didShrink && LOG.isDebugEnabled()) {
            LOG.debug("Cellule de tableau retrecie a {}% (hauteur texte residuelle {}pt -> ligne {}pt) "
                            + "apres correction d'interligne",
                    Math.round(factor * 100), textHeight, anchor.getHeight());
        }
        return didShrink;
    }

    /**
     * Cellules du tableau portant au moins un caractere de texte visible, chacune une
     * seule fois (une cellule fusionnee peut apparaitre a plusieurs positions {@code
     * (row,col)} de la grille - voir {@link XSLFTableRow#getCells()} - dedoublonnee par
     * identite d'objet pour ne jamais lui appliquer deux fois la correction d'interligne,
     * qui n'est pas idempotente).
     */
    static List<XSLFTableCell> collectCellsWithVisibleText(XSLFTable table) {
        Set<XSLFTableCell> seen = Collections.newSetFromMap(new IdentityHashMap<>());
        List<XSLFTableCell> result = new ArrayList<>();
        for (XSLFTableRow row : table.getRows()) {
            for (XSLFTableCell cell : row.getCells()) {
                if (cell == null || !seen.add(cell)) {
                    continue;
                }
                String text = cell.getText();
                if (text != null && !text.trim().isEmpty()) {
                    result.add(cell);
                }
            }
        }
        return result;
    }

    /** Parcourt les formes du slide, y compris a l'interieur des groupes, et ne garde que les tableaux. */
    static List<XSLFTable> collectTables(List<XSLFShape> shapes) {
        List<XSLFTable> result = new ArrayList<>();
        for (XSLFShape shape : shapes) {
            if (shape instanceof XSLFGroupShape) {
                result.addAll(collectTables(((XSLFGroupShape) shape).getShapes()));
            } else if (shape instanceof XSLFTable) {
                result.add((XSLFTable) shape);
            }
        }
        return result;
    }
}
