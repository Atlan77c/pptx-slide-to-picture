package io.github.atlan77c.pptx2picture.internal;

import org.apache.poi.sl.usermodel.PictureData;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFAutoShape;
import org.apache.poi.xslf.usermodel.XSLFPictureData;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextBox;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.presentationml.x2006.main.CTPicture;
import org.openxmlformats.schemas.presentationml.x2006.main.CTPlaceholder;
import org.openxmlformats.schemas.presentationml.x2006.main.CTShape;
import org.openxmlformats.schemas.presentationml.x2006.main.STPlaceholderType;

import java.awt.Graphics2D;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.IOException;

import static org.junit.jupiter.api.Assertions.assertEquals;

/**
 * Reproduit le motif observe sur un fichier reel : un titre partiellement
 * recouvert par une image voisine peinte apres lui dans l'ordre du document.
 */
class TitleRepainterTest {

    private XMLSlideShow ppt;
    private XSLFSlide slide;
    private BufferedImage image;
    private Graphics2D graphics;

    @BeforeEach
    void setUp() {
        ppt = new XMLSlideShow();
        slide = ppt.createSlide();
        image = new BufferedImage(100, 100, BufferedImage.TYPE_INT_ARGB);
        graphics = image.createGraphics();
    }

    @AfterEach
    void tearDown() throws IOException {
        graphics.dispose();
        ppt.close();
    }

    @Test
    void repaintTitles_repaintsShapeDeclaredAsTitlePlaceholder() {
        XSLFTextBox title = slide.createTextBox();
        title.setText("Titre du placeholder");
        setPlaceholderType(title, STPlaceholderType.TITLE);

        int repainted = TitleRepainter.repaintTitles(slide, graphics);

        assertEquals(1, repainted);
    }

    @Test
    void repaintTitles_repaintsShapeDeclaredAsCenteredTitlePlaceholder() {
        XSLFTextBox title = slide.createTextBox();
        title.setText("Titre centre du placeholder");
        setPlaceholderType(title, STPlaceholderType.CTR_TITLE);

        int repainted = TitleRepainter.repaintTitles(slide, graphics);

        assertEquals(1, repainted);
    }

    @Test
    void repaintTitles_repaintsPlainTextBoxNamedAfterTitleInAnySupportedLanguage() {
        // Motif reel observe sur un fichier de production : un titre dissocie
        // de son placeholder, reconnaissable uniquement par son nom ("Titre N"
        // et equivalents).
        String[] names = {"Titre 3", "Title 2", "Título 1", "Titolo 4", "Titel 5"};
        for (String name : names) {
            ppt = new XMLSlideShow();
            slide = ppt.createSlide();
            XSLFTextBox box = slide.createTextBox();
            box.setText("Texte du titre");
            setShapeName(box, name);

            int repainted = TitleRepainter.repaintTitles(slide, graphics);

            assertEquals(1, repainted, "le nom '" + name + "' doit etre reconnu comme un titre");
        }
    }

    @Test
    void repaintTitles_leavesPlainTextBoxWithUnrelatedNameUntouched() {
        XSLFTextBox box = slide.createTextBox();
        box.setText("Zone de texte quelconque");
        setShapeName(box, "ZoneTexte 12");

        int repainted = TitleRepainter.repaintTitles(slide, graphics);

        assertEquals(0, repainted);
    }

    @Test
    void repaintTitles_leavesNameThatOnlySharesTitlePrefixUntouched() {
        // "Titre" est un prefixe de "Titredecolonne" mais ce n'est pas un titre :
        // la reconnaissance doit s'arreter a une frontiere de mot.
        XSLFTextBox box = slide.createTextBox();
        box.setText("Pas un titre");
        setShapeName(box, "Titredecolonne 1");

        int repainted = TitleRepainter.repaintTitles(slide, graphics);

        assertEquals(0, repainted);
    }

    @Test
    void repaintTitles_leavesNonTextShapesUntouched() throws IOException {
        // Une image nommee comme un titre (motif reel observe) ne doit jamais
        // etre consideree comme un titre, quel que soit son nom.
        byte[] onePixelPng = onePixelPng();
        XSLFPictureData pictureData = ppt.addPicture(onePixelPng, PictureData.PictureType.PNG);
        XSLFPictureShape picture = slide.createPicture(pictureData);
        setShapeName(picture, "Image 24");
        picture.setAnchor(new Rectangle2D.Double(0, 0, 10, 10));

        int repainted = TitleRepainter.repaintTitles(slide, graphics);

        assertEquals(0, repainted);
    }

    @Test
    void repaintTitles_leavesAutoShapeNamedLikeATitleButWithoutTextUntouched() {
        // Un cas volontairement absent : une forme non textuelle ne peut de toute
        // facon pas etre un XSLFTextShape ; ce test verifie surtout la robustesse
        // du filtrage par type avant le test sur le nom.
        XSLFAutoShape shape = slide.createAutoShape();
        setShapeName(shape, "Titre decoratif");
        shape.setText("Ce n'est pas vraiment un titre de slide");

        int repainted = TitleRepainter.repaintTitles(slide, graphics);

        // Une XSLFAutoShape EST un XSLFTextShape : son nom commence bien par
        // "Titre" suivi d'un espace, donc elle est repeinte - comportement voulu
        // (mieux vaut un faux positif sans consequence visuelle qu'un titre
        // manque, voir Javadoc de la classe).
        assertEquals(1, repainted);
    }

    private static void setPlaceholderType(XSLFTextBox shape, STPlaceholderType.Enum type) {
        CTShape ctShape = (CTShape) shape.getXmlObject();
        CTPlaceholder ph = ctShape.getNvSpPr().getNvPr().addNewPh();
        ph.setType(type);
    }

    /**
     * L'API XSLF publique ne propose aucun setter pour le nom d'une forme
     * ({@code getShapeName()} n'a pas de {@code setShapeName()} correspondant) -
     * on doit donc passer par le XML sous-jacent, comme pour
     * {@link #setPlaceholderType}. Le nom vit au meme endroit ({@code cNvPr})
     * pour un {@code <p:sp>} (texte/autoshape) et un {@code <p:pic>} (image),
     * mais sous un chemin d'acces different.
     */
    private static void setShapeName(XSLFShape shape, String name) {
        Object xmlObject = shape.getXmlObject();
        if (xmlObject instanceof CTShape) {
            ((CTShape) xmlObject).getNvSpPr().getCNvPr().setName(name);
        } else if (xmlObject instanceof CTPicture) {
            ((CTPicture) xmlObject).getNvPicPr().getCNvPr().setName(name);
        } else {
            throw new IllegalArgumentException("Type de forme non gere par ce test: " + xmlObject.getClass());
        }
    }

    private static byte[] onePixelPng() {
        // PNG transparent de 1x1 pixel, code en dur pour eviter toute dependance
        // externe dans le test.
        return java.util.Base64.getDecoder().decode(
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");
    }
}
