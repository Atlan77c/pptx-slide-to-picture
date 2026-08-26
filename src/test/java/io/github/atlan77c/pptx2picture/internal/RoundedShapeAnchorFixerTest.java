package io.github.atlan77c.pptx2picture.internal;

import org.apache.poi.sl.usermodel.ShapeType;
import org.apache.poi.sl.usermodel.TextShape;
import org.apache.poi.sl.usermodel.VerticalAlignment;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFAutoShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextBox;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextBodyProperties;
import org.openxmlformats.schemas.presentationml.x2006.main.CTShape;

import java.io.IOException;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * Reproduit le motif observe sur un fichier reel : une zone de texte
 * {@code spAutoFit} dont le contour a ete change en ellipse (via "Modifier la
 * forme" dans PowerPoint), sans ancrage vertical explicite - le texte
 * apparaissait colle en haut de la forme au lieu d'etre centre.
 */
class RoundedShapeAnchorFixerTest {

    private XMLSlideShow ppt;
    private XSLFSlide slide;

    @BeforeEach
    void setUp() {
        ppt = new XMLSlideShow();
        slide = ppt.createSlide();
    }

    @AfterEach
    void tearDown() throws IOException {
        ppt.close();
    }

    @Test
    void fixVerticalAnchor_forcesMiddle_forNonRectangularSpAutoFitShapeWithoutExplicitAnchor() {
        XSLFAutoShape ellipse = slide.createAutoShape();
        ellipse.setShapeType(ShapeType.ELLIPSE);
        ellipse.setText("Texte dans une zone redimensionnee en ellipse");
        ellipse.setTextAutofit(TextShape.TextAutofit.SHAPE);

        // Preconditions verifiees explicitement (plutot que suppposees) : un fichier
        // reel omet l'attribut "anchor", mais rien ne garantit que le bodyPr genere
        // par l'API de creation de forme de POI parte lui aussi sans "anchor" explicite
        // - on le retire donc nous-memes au niveau XML pour reproduire fidelement le
        // motif reel (voir Javadoc de la classe), independamment de ce detail
        // d'implementation de la fabrique de formes de POI.
        assertEquals(ShapeType.ELLIPSE, ellipse.getShapeType(), "precondition : geometrie non rectangulaire");
        assertEquals(TextShape.TextAutofit.SHAPE, ellipse.getTextAutofit(), "precondition : autofit spAutoFit");
        Object xmlObject = ellipse.getXmlObject();
        assertTrue(xmlObject instanceof CTShape, "precondition : forme adossee a un CTShape (p:sp)");
        CTTextBodyProperties bodyPr = ((CTShape) xmlObject).getTxBody().getBodyPr();
        if (bodyPr.isSetAnchor()) {
            bodyPr.unsetAnchor();
        }
        assertFalse(bodyPr.isSetAnchor(), "precondition : aucun ancrage vertical explicite (motif observe sur un fichier reel)");

        int fixed = RoundedShapeAnchorFixer.fixVerticalAnchor(slide);

        assertEquals(1, fixed);
        assertEquals(VerticalAlignment.MIDDLE, ellipse.getVerticalAlignment());
    }

    @Test
    void fixVerticalAnchor_leavesRectangularShapesUntouched() {
        XSLFAutoShape rect = slide.createAutoShape();
        rect.setShapeType(ShapeType.RECT);
        rect.setText("Texte dans un rectangle spAutoFit");
        rect.setTextAutofit(TextShape.TextAutofit.SHAPE);

        int fixed = RoundedShapeAnchorFixer.fixVerticalAnchor(slide);

        assertEquals(0, fixed, "un rectangle n'a pas de marge liee a la courbure - rien a corriger");
    }

    @Test
    void fixVerticalAnchor_leavesNonSpAutoFitShapesUntouched() {
        XSLFAutoShape ellipse = slide.createAutoShape();
        ellipse.setShapeType(ShapeType.ELLIPSE);
        ellipse.setText("Texte dans une ellipse a taille fixe");
        ellipse.setTextAutofit(TextShape.TextAutofit.NONE);

        int fixed = RoundedShapeAnchorFixer.fixVerticalAnchor(slide);

        assertEquals(0, fixed, "le mecanisme suspecte (marge spAutoFit liee a la courbure) ne concerne "
                + "que les formes qui se redimensionnent elles-memes selon leur texte");
    }

    @Test
    void fixVerticalAnchor_neverOverridesAnExplicitlyDeclaredAnchor() {
        XSLFAutoShape ellipse = slide.createAutoShape();
        ellipse.setShapeType(ShapeType.ELLIPSE);
        ellipse.setText("Ancrage choisi explicitement par l'auteur");
        ellipse.setTextAutofit(TextShape.TextAutofit.SHAPE);
        ellipse.setVerticalAlignment(VerticalAlignment.TOP); // choix explicite, meme s'il coincide avec le defaut

        int fixed = RoundedShapeAnchorFixer.fixVerticalAnchor(slide);

        assertEquals(0, fixed, "un anchor explicitement present dans le bodyPr ne doit jamais etre ecrase");
        assertEquals(VerticalAlignment.TOP, ellipse.getVerticalAlignment());
    }

    @Test
    void fixVerticalAnchor_leavesPlainTextBoxesUntouched() {
        // Une XSLFTextBox "classique" (jamais changee en forme non rectangulaire cote
        // PowerPoint) n'a pas de prstGeom explicite - getShapeType() doit alors etre
        // traite comme "rectangulaire" (pas de marge de courbure a compenser).
        XSLFTextBox box = slide.createTextBox();
        box.setText("Zone de texte classique");
        box.setTextAutofit(TextShape.TextAutofit.SHAPE);

        int fixed = RoundedShapeAnchorFixer.fixVerticalAnchor(slide);

        assertEquals(0, fixed);
    }
}
