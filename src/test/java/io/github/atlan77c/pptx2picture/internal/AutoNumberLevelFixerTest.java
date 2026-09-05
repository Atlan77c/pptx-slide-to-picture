package io.github.atlan77c.pptx2picture.internal;

import org.apache.poi.sl.usermodel.AutoNumberingScheme;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextBox;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextParagraphProperties;

import java.awt.Graphics2D;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * Tests unitaires pour {@link AutoNumberLevelFixer}.
 *
 * <p>Reproduit le bug d'origine (voir Javadoc de {@link AutoNumberLevelFixer}) :
 * un compteur de numerotation automatique PLAT partage par tous les
 * paragraphes d'une forme, sans distinction de niveau d'indentation. La
 * majorite des cas sont testes directement sur {@link
 * AutoNumberLevelFixer.LevelNumbering} (aucune metrique de police/rendu
 * requise, comportement purement deterministe) - meme discipline que les
 * tests existants de ce projet qui testent en priorite la logique pure
 * plutot que le rendu pixel par pixel.
 */
class AutoNumberLevelFixerTest {

    // ------------------------------------------------------------------
    // LevelNumbering : logique pure, sans POI ni Graphics2D.
    // ------------------------------------------------------------------

    @Test
    void levelNumbering_singleLevel_incrementsSequentially() {
        AutoNumberLevelFixer.LevelNumbering numbering = new AutoNumberLevelFixer.LevelNumbering();

        assertEquals(1, numbering.next(0, null));
        assertEquals(2, numbering.next(0, null));
        assertEquals(3, numbering.next(0, null));
    }

    @Test
    void levelNumbering_deeperLevel_restartsAt1AndParentContinuesAfterwards() {
        // Reproduit exactement la structure du sommaire reel (slide 5) : chaque item de
        // niveau 0 est suivi d'items de niveau 1 qui doivent redemarrer independamment.
        AutoNumberLevelFixer.LevelNumbering numbering = new AutoNumberLevelFixer.LevelNumbering();

        assertEquals(1, numbering.next(0, null), "Section A");
        assertEquals(1, numbering.next(1, null), "Section A, sous-point 1");
        assertEquals(2, numbering.next(1, null), "Section A, sous-point 2");
        assertEquals(3, numbering.next(1, null), "Section A, sous-point 3");

        assertEquals(2, numbering.next(0, null), "Definir - NE DOIT PAS etre '1.' comme avec le "
                + "compteur plat de POI, ni continuer la sequence a '5.'");
        assertEquals(1, numbering.next(1, null), "Section B, sous-point 1 - DOIT redemarrer a 1, "
                + "pas continuer a partir de 4");
        assertEquals(2, numbering.next(1, null));
        assertEquals(3, numbering.next(1, null));

        assertEquals(3, numbering.next(0, null), "Section C");
        assertEquals(1, numbering.next(1, null), "Section C, sous-point 1 - redemarre encore");
    }

    @Test
    void levelNumbering_explicitStartAt_becomesNewBaselineForFollowingSiblingsAndChildren() {
        // Reproduit le slide de test de l'utilisateur : buAutoNum/@startAt="4" pose sur un
        // item de niveau 0 ("Section B"), verifie que (a) l'item lui-meme prend bien 4, (b) les
        // enfants de niveau 1 en dessous redemarrent a 1 SANS etre contamines par ce 4 - le
        // coeur du bug d'origine (POI les faisait continuer a 5, 6, 7...).
        AutoNumberLevelFixer.LevelNumbering numbering = new AutoNumberLevelFixer.LevelNumbering();

        assertEquals(1, numbering.next(0, null), "Section A");
        assertEquals(1, numbering.next(1, null));
        assertEquals(2, numbering.next(1, null));
        assertEquals(3, numbering.next(1, null));

        assertEquals(4, numbering.next(0, 4), "Definir, startAt=4 explicite");
        assertEquals(1, numbering.next(1, null), "Section B, sous-point 1 - NE DOIT PAS valoir 5");
        assertEquals(2, numbering.next(1, null));
        assertEquals(3, numbering.next(1, null));
        assertEquals(4, numbering.next(1, null));

        // startAt n'est qu'un PLANCHER (voir levelNumbering_secondExplicitStartAt_isFlooredBy...
        // ci-dessous pour le cas reel qui a revele ce point) : la continuation naturelle du
        // niveau 0 (4 -> 5) est ICI DEJA au-dessus de startAt=4, donc startAt n'a aucun effet -
        // "Section C" doit valoir 5, PAS 4 (un ecrasement inconditionnel par startAt, comme
        // dans une version anterieure de ce correctif, produirait a tort 4 ici aussi).
        assertEquals(5, numbering.next(0, 4), "La trajectoire, startAt=4 explicite mais deja "
                + "depasse par la continuation naturelle du niveau 0 (4 -> 5)");
        assertEquals(1, numbering.next(1, null), "Section C, sous-point 1 - redemarre a 1");
    }

    @Test
    void levelNumbering_secondExplicitStartAt_isFlooredByNaturalContinuation_notHardReset() {
        // Regression reelle (signalee par l'utilisateur le 2026-09-05) : rejoue exactement les
        // 18 paragraphes du slide 6 de "fichier-test-B.pptx"
        // (lvl/startAt extraits directement du XML). Attendu par PowerPoint : la sequence de
        // niveau 0 est 1, 4, 5 - PAS 1, 4, 4. Une premiere version de ce correctif, qui ecrasait
        // inconditionnellement la valeur par startAt des qu'il etait present (au lieu de le
        // traiter comme un plancher sur la continuation naturelle), produisait a tort 1, 4, 4 :
        // le second startAt="4" (sur "Section C") ecrasait la continuation naturelle du
        // niveau 0, qui valait deja 5 a ce point (puisque "Section B" avait deja pris la valeur 4).
        AutoNumberLevelFixer.LevelNumbering numbering = new AutoNumberLevelFixer.LevelNumbering();

        assertEquals(1, numbering.next(0, null), "Section A");
        assertEquals(1, numbering.next(1, null), "Section A, sous-point 1");
        assertEquals(2, numbering.next(1, null), "Section A, sous-point 2");
        assertEquals(3, numbering.next(1, null), "Section A, sous-point 3");
        // paragraphe-espaceur (buAutoNum absent) : jamais transmis a next(), voir plus bas.

        assertEquals(4, numbering.next(0, 4), "Definir (startAt=4 explicite dans le XML)");
        assertEquals(1, numbering.next(1, null), "Section B, sous-point 1");
        assertEquals(2, numbering.next(1, null), "Section B, sous-point 2");
        assertEquals(3, numbering.next(1, null), "Section B, sous-point 3");
        assertEquals(4, numbering.next(1, null), "Section B, sous-point 4");
        // paragraphe-espaceur : jamais transmis a next().

        assertEquals(5, numbering.next(0, 4), "La trajectoire (startAt=4 explicite lui aussi dans "
                + "le XML, mais NE DOIT PAS regresser la sequence de niveau 0 qui est deja a 5)");
        assertEquals(1, numbering.next(1, null), "Section C, sous-point 1");
        assertEquals(2, numbering.next(1, null), "Section C, sous-point 2");
        assertEquals(3, numbering.next(1, null), "Section C, sous-point 3");
        assertEquals(4, numbering.next(1, null), "Section C, sous-point 4");
        // 2 paragraphes-espaceurs finaux : jamais transmis a next().
    }

    @Test
    void levelNumbering_returningToShallowerLevel_resumesItsOwnSequence_notRestarts() {
        // Cas a 3 niveaux : redescendre d'un niveau profond vers un niveau deja visite doit
        // REPRENDRE sa sequence (pas la redemarrer), alors que re-descendre ensuite dans le
        // niveau profond doit lui, a nouveau, redemarrer (nouveau sous-groupe).
        AutoNumberLevelFixer.LevelNumbering numbering = new AutoNumberLevelFixer.LevelNumbering();

        assertEquals(1, numbering.next(0, null), "A (lvl0)");
        assertEquals(1, numbering.next(1, null), "x (lvl1)");
        assertEquals(1, numbering.next(2, null), "p (lvl2)");
        assertEquals(2, numbering.next(2, null), "q (lvl2)");
        assertEquals(2, numbering.next(1, null), "y (lvl1) - reprend x=1, ne redemarre pas a 1");
        assertEquals(1, numbering.next(2, null), "r (lvl2) - redemarre : nouveau parent lvl1");
    }

    @Test
    void levelNumbering_nonNumberedParagraph_isSimplyNeverPassed_doesNotDisturbState() {
        // Un paragraphe sans puce automatique (espaceur) n'appelle jamais next() en
        // production (voir LevelAwareTextShape.drawParagraphs) - verifie que l'etat reste
        // intact autour d'un tel "trou" dans la sequence des appels.
        AutoNumberLevelFixer.LevelNumbering numbering = new AutoNumberLevelFixer.LevelNumbering();

        assertEquals(1, numbering.next(1, null));
        assertEquals(2, numbering.next(1, null));
        // ... un paragraphe vide entre les deux, jamais transmis a next() ...
        assertEquals(3, numbering.next(1, null), "la sequence continue normalement apres le trou");
    }

    // ------------------------------------------------------------------
    // Extraction depuis de vrais XSLFTextParagraph (getIndentLevel/BulletStyle),
    // sans rendu Graphics2D - verifie le cablage entre l'API POI reelle et
    // LevelNumbering, independamment du rendu proprement dit.
    // ------------------------------------------------------------------

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
    void realParagraphs_summaryLikeStructure_producesLevelAwareSequence() {
        XSLFTextBox box = slide.createTextBox();
        box.clearText(); // etat de depart garanti : zero paragraphe, voir addNumberedParagraph/addSpacerParagraph

        addNumberedParagraph(box, 0, "Section A", null);
        addNumberedParagraph(box, 1, "Section A, sous-point 1", null);
        addNumberedParagraph(box, 1, "Section A, sous-point 2", null);
        addSpacerParagraph(box);
        addNumberedParagraph(box, 0, "Section B", null);
        addNumberedParagraph(box, 1, "Section B, sous-point 1", null);
        addNumberedParagraph(box, 1, "Section B, sous-point 2", null);

        List<Integer> rendered = new ArrayList<>();
        AutoNumberLevelFixer.LevelNumbering numbering = new AutoNumberLevelFixer.LevelNumbering();
        for (XSLFTextParagraph p : box.getTextParagraphs()) {
            if (p.getBulletStyle() == null || p.getBulletStyle().getAutoNumberingScheme() == null) {
                continue; // espaceur : jamais transmis a next(), voir LevelAwareTextShape.drawParagraphs
            }
            rendered.add(numbering.next(p.getIndentLevel(), p.getBulletStyle().getAutoNumberingStartAt()));
        }

        assertEquals(List.of(1, 1, 2, 2, 1, 2), rendered,
                "Comprendre=1 ; ses 2 enfants=1,2 ; Definir=2 (pas '1.' comme avec POI standard) ; "
                        + "ses 2 enfants redemarrent a 1,2 (pas '3.','4.')");
    }

    // ------------------------------------------------------------------
    // Rendu reel de bout en bout : verifie que l'appel reflexif a
    // DrawTextParagraph#breakText/#setFirstParagraph fonctionne effectivement
    // contre la version d'Apache POI presente sur le classpath (echec sinon,
    // voir Javadoc de la classe, section "Pourquoi de la reflexion").
    // ------------------------------------------------------------------

    @Test
    void installBeforeDraw_thenSlideDraw_rendersLevelAwareSummaryWithoutThrowing() {
        XSLFTextBox box = slide.createTextBox();
        box.setAnchor(new Rectangle2D.Double(50, 50, 400, 300));
        box.clearText(); // etat de depart garanti : zero paragraphe, voir addNumberedParagraph/addSpacerParagraph
        addNumberedParagraph(box, 0, "Section A", null);
        addNumberedParagraph(box, 1, "Section A, sous-point 1", null);
        addNumberedParagraph(box, 1, "Section A, sous-point 2", null);
        addSpacerParagraph(box);
        addNumberedParagraph(box, 0, "Section B", 4);
        addNumberedParagraph(box, 1, "Section B, sous-point 1", null);

        BufferedImage img = new BufferedImage(600, 600, BufferedImage.TYPE_INT_ARGB);
        Graphics2D graphics = img.createGraphics();
        try {
            Object previous = AutoNumberLevelFixer.installBeforeDraw(graphics);
            try {
                assertDoesNotThrow(() -> slide.draw(graphics),
                        "l'appel reflexif a DrawTextParagraph#breakText/#setFirstParagraph doit reussir "
                                + "contre la version reelle d'Apache POI du classpath");

                // getTextHeight() passe par le MEME DrawFactory que celui installe sur `graphics`
                // (voir XSLFTextShape#getTextHeight(Graphics2D) -> DrawFactory.getInstance(graphics))
                // - verifie donc, elle aussi, le chemin corrige plutot que le DrawTextShape standard.
                double height1 = box.getTextHeight(graphics);
                double height2 = box.getTextHeight(graphics);
                assertTrue(height1 > 0, "la hauteur mesuree doit etre positive pour un texte non vide");
                assertEquals(height1, height2, 0.001,
                        "la mesure doit etre deterministe (LevelNumbering repart de zero a chaque appel)");
            } finally {
                AutoNumberLevelFixer.restoreAfterDraw(graphics, previous);
            }
        } finally {
            graphics.dispose();
        }
    }

    private static void addNumberedParagraph(XSLFTextBox box, int level, String text, Integer startAt) {
        XSLFTextParagraph para = box.addNewTextParagraph();
        para.setIndentLevel(level);
        para.addNewTextRun().setText(text);
        // XSLFTextParagraph#isBullet() (donc #getBulletStyle() != null) ne regarde QUE buFont/
        // buChar/buNone, jamais buAutoNum directement - dans un fichier PowerPoint reel c'est
        // toujours vrai par heritage du theme/masque (styles de liste par defaut), mais un
        // XMLSlideShow fraichement cree par ce test n'a aucun style de liste herite : poser
        // explicitement setBullet(true) avant setBulletAutoNumber(...) le rend independant de
        // cette hypothese d'heritage.
        para.setBullet(true);
        if (startAt != null) {
            para.setBulletAutoNumber(AutoNumberingScheme.arabicPeriod, startAt);
        } else {
            // Reproduit le cas reel le plus courant : buAutoNum SANS @startAt (poursuite de la
            // sequence du niveau). setBulletAutoNumber() de POI impose toujours un startAt >= 1 -
            // on le pose puis on le retire explicitement de l'XML sous-jacent (meme technique que
            // POI utilise en interne pour buFont/buChar/buAutoNum - isSetX()/unsetX()).
            para.setBulletAutoNumber(AutoNumberingScheme.arabicPeriod, 1);
            CTTextParagraphProperties pr = para.getXmlObject().getPPr();
            if (pr != null && pr.isSetBuAutoNum() && pr.getBuAutoNum().isSetStartAt()) {
                pr.getBuAutoNum().unsetStartAt();
            }
        }
    }

    private static void addSpacerParagraph(XSLFTextBox box) {
        XSLFTextParagraph para = box.addNewTextParagraph();
        para.addNewTextRun().setText(" ");
        // pas de setBulletAutoNumber(...) : aucune puce automatique, exactement comme les
        // paragraphes-espaceurs des fichiers reels examines.
    }
}
