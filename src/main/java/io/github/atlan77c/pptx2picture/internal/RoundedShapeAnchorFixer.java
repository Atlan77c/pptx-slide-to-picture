package io.github.atlan77c.pptx2picture.internal;

import org.apache.poi.sl.usermodel.ShapeType;
import org.apache.poi.sl.usermodel.TextShape;
import org.apache.poi.sl.usermodel.VerticalAlignment;
import org.apache.poi.xslf.usermodel.XSLFGroupShape;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTable;
import org.apache.poi.xslf.usermodel.XSLFTableCell;
import org.apache.poi.xslf.usermodel.XSLFTableRow;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextBody;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextBodyProperties;
import org.openxmlformats.schemas.presentationml.x2006.main.CTShape;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.ArrayList;
import java.util.List;

/**
 * Corrige un ecart de rendu decouvert sur un vrai fichier (slide 30) : du
 * texte affiche colle en haut d'une forme au lieu d'etre centre
 * verticalement, alors que la forme est bien {@code spAutoFit}
 * ("redimensionner la forme selon le texte").
 *
 * <p><b>Motif observe</b> : une zone de texte (souvent creee via Insertion
 * &gt; Zone de texte, {@code cNvSpPr txBox="1"}) a vu son contour change en
 * une geometrie non rectangulaire (ex. {@code prstGeom prst="ellipse"}, via
 * "Modifier la forme" dans PowerPoint) - deux exemples reels : {@code ZoneTexte 160}
 * ("Affectation restreinte au conseiller TH") et {@code ZoneTexte 195} ("Pas
 * d'affectation autre que FT ou CapE"). Son {@code bodyPr} ne declare aucun
 * {@code anchor} explicite - valeur par defaut OOXML "Haut", normalement sans
 * consequence visuelle pour une forme {@code spAutoFit} puisque PowerPoint
 * redimensionne la boite pour epouser exactement le texte (aucun espace
 * residuel, donc Haut ou Milieu produisent le meme rendu).
 *
 * <p><b>Hypothese retenue</b> (coherente avec l'absence totale de ces formes
 * dans les logs de {@link OverflowAwareTextFitter} - jamais de debordement
 * mesure dessus) : pour une geometrie non rectangulaire, PowerPoint reserve,
 * lors du calcul {@code spAutoFit}, une marge supplementaire liee a la
 * courbure de la forme (le texte doit rester dans la zone utile de
 * l'ellipse, plus etroite pres du haut/bas que sur un rectangle) - la boite
 * stockee dans le fichier est donc plus haute que ce qu'exigerait une mise
 * en page purement rectangulaire du meme texte. Apache POI, lui, mesure et
 * positionne le texte comme s'il s'agissait d'un simple rectangle (aucune
 * prise en compte de la geometrie pour le texte), sans jamais detecter de
 * "debordement" puisque le texte tient dans cette boite deja plus haute que
 * necessaire. L'espace residuel, avec un ancrage "Haut" par defaut, se
 * retrouve donc entierement en bas de la forme - visuellement, le texte
 * parait pousse vers le haut.
 *
 * <p><b>Correctif</b> : pour toute forme de texte {@code spAutoFit} dont la
 * geometrie declaree n'est pas un simple rectangle et dont le {@code bodyPr}
 * ne declare aucun {@code anchor} explicite (on ne touche jamais a un choix
 * explicite de l'auteur), l'ancrage vertical est force a {@code MIDDLE} -
 * l'espace residuel, quelle qu'en soit l'origine exacte, est ainsi reparti
 * de part et d'autre du texte plutot que laisse entierement en bas.
 */
public final class RoundedShapeAnchorFixer {

    private static final Logger LOG = LoggerFactory.getLogger(RoundedShapeAnchorFixer.class);

    private RoundedShapeAnchorFixer() {
    }

    /**
     * Corrige, in-place, l'ancrage vertical des formes concernees du slide.
     * A appeler avant {@code slide.draw(graphics)}, et avant
     * {@link OverflowAwareTextFitter#fitOverflowingText} pour que ce dernier
     * calcule ses zones de debordement avec l'ancrage deja corrige.
     *
     * @return le nombre de formes effectivement corrigees.
     */
    public static int fixVerticalAnchor(XSLFSlide slide) {
        int fixed = 0;
        for (XSLFShape shape : collectTextShapes(slide.getShapes())) {
            XSLFTextShape ts = (XSLFTextShape) shape;
            if (shouldForceMiddleAnchor(ts)) {
                ts.setVerticalAlignment(VerticalAlignment.MIDDLE);
                fixed++;
                if (LOG.isDebugEnabled()) {
                    LOG.debug("{} : forme spAutoFit non rectangulaire ({}) sans ancrage vertical declare -> "
                            + "ancrage force a MIDDLE (voir Javadoc de la classe)",
                            shape.getShapeName(), ts.getShapeType());
                }
            }
        }
        return fixed;
    }

    private static boolean shouldForceMiddleAnchor(XSLFTextShape ts) {
        if (ts.getTextAutofit() != TextShape.TextAutofit.SHAPE) {
            return false;
        }
        ShapeType type = ts.getShapeType();
        if (type == null || type == ShapeType.RECT) {
            return false;
        }
        return !hasExplicitVerticalAnchor(ts);
    }

    /**
     * {@code true} si le {@code bodyPr} de la forme declare explicitement un
     * attribut {@code anchor} (choix delibere de l'auteur, jamais ecrase) -
     * {@code false} s'il est absent (valeur par defaut OOXML implicite) ou si
     * la structure XML sous-jacente n'est pas celle attendue (prudence : dans
     * le doute, ne pas toucher au comportement par defaut).
     */
    private static boolean hasExplicitVerticalAnchor(XSLFTextShape ts) {
        Object xmlObject = ts.getXmlObject();
        if (!(xmlObject instanceof CTShape)) {
            return true;
        }
        CTShape ctShape = (CTShape) xmlObject;
        if (!ctShape.isSetTxBody()) {
            return true;
        }
        // bodyPr est un element obligatoire de txBody (contrairement a txBody lui-meme,
        // optionnel) : XMLBeans ne genere donc pas de isSetBodyPr() pour lui, seulement
        // pour ses propres attributs optionnels comme "anchor" ci-dessous.
        CTTextBody txBody = ctShape.getTxBody();
        CTTextBodyProperties bodyPr = txBody.getBodyPr();
        return bodyPr == null || bodyPr.isSetAnchor();
    }

    /** Parcourt les formes du slide, y compris a l'interieur des groupes ET des cellules de tableau (voir {@link SymbolFontRunFixer}). */
    private static List<XSLFShape> collectTextShapes(List<XSLFShape> shapes) {
        List<XSLFShape> result = new ArrayList<>();
        for (XSLFShape shape : shapes) {
            if (shape instanceof XSLFGroupShape) {
                result.addAll(collectTextShapes(((XSLFGroupShape) shape).getShapes()));
            } else if (shape instanceof XSLFTable) {
                for (XSLFTableRow row : ((XSLFTable) shape).getRows()) {
                    for (XSLFTableCell cell : row.getCells()) {
                        result.add(cell);
                    }
                }
            } else if (shape instanceof XSLFTextShape) {
                result.add(shape);
            }
        }
        return result;
    }
}
