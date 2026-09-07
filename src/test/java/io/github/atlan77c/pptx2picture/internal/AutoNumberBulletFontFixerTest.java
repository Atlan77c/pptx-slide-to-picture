package io.github.atlan77c.pptx2picture.internal;

import org.apache.poi.sl.usermodel.AutoNumberingScheme;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFSlideMaster;
import org.apache.poi.xslf.usermodel.XSLFTextBox;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextListStyle;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextParagraphProperties;
import org.openxmlformats.schemas.presentationml.x2006.main.CTSlideMasterTextStyles;

import java.io.IOException;

import static org.junit.jupiter.api.Assertions.assertEquals;

/**
 * Tests unitaires pour {@link AutoNumberBulletFontFixer}.
 *
 * <p>Reproduit le bug d'origine (voir Javadoc de la classe, et de {@link
 * AutoNumberLevelFixer} pour le renvoi croise) sur le fichier reel qui l'a
 * revele (slide 96 d'un document interne reel) : une liste NUMEROTEE automatiquement
 * (buAutoNum) herite, via {@code XSLFTextParagraph#getBulletFont()}, une
 * police Wingdings/Symbol posee ailleurs dans la chaine pour un tout autre
 * choix de puce (un buChar litteral) - jamais localement sur le paragraphe
 * numerote lui-meme.
 *
 * <p>Pour reproduire fidelement cet heritage (par opposition a poser
 * directement {@code buFont} SUR le paragraphe, ce qui neutraliserait le
 * symptome que ce correctif cible - voir sa garde "pas de buFont local"),
 * ces tests injectent le buFont/buChar Wingdings dans le style de texte par
 * defaut du MASQUE de diapositive ({@code txStyles/otherStyle/lvl1pPr}, le
 * chemin reellement emprunte par {@code XSLFTextParagraph#getDefaultMasterStyle()}
 * pour une zone de texte libre - PAS un placeholder), exactement comme le
 * layout du fichier reel le fait pour son placeholder "corps".
 */
class AutoNumberBulletFontFixerTest {

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

    /**
     * Pose {@code <a:buFont typeface="..."/><a:buChar char="..."/>} sur le
     * {@code lvl1pPr} du style "otherStyle" du masque - le style dont herite
     * une zone de texte libre (pas un placeholder) de niveau d'indentation 0,
     * voir Javadoc de la classe.
     */
    private void setMasterOtherStyleLevel1Bullet(String typeface, String bulletChar) {
        XSLFSlideMaster master = ppt.getSlideMasters().get(0);
        CTSlideMasterTextStyles txStyles = master.getXmlObject().isSetTxStyles()
                ? master.getXmlObject().getTxStyles() : master.getXmlObject().addNewTxStyles();
        CTTextListStyle otherStyle =
                txStyles.isSetOtherStyle() ? txStyles.getOtherStyle() : txStyles.addNewOtherStyle();
        CTTextParagraphProperties lvl1 = otherStyle.isSetLvl1PPr() ? otherStyle.getLvl1PPr() : otherStyle.addNewLvl1PPr();
        (lvl1.isSetBuFont() ? lvl1.getBuFont() : lvl1.addNewBuFont()).setTypeface(typeface);
        (lvl1.isSetBuChar() ? lvl1.getBuChar() : lvl1.addNewBuChar()).setChar(bulletChar);
    }

    @Test
    void fixAutoNumberBulletFonts_correctsFontInheritedFromMasterCharBulletDefault() {
        // Reproduit exactement le slide 96 reel : le layout definit, pour le placeholder
        // "corps", un buFont Wingdings/buChar par defaut (destine a une puce a caractere
        // litteral) ; nos paragraphes numerotes n'ont AUCUN buFont local et heritent donc,
        // a tort, de ce Wingdings pour dessiner leur numero.
        setMasterOtherStyleLevel1Bullet("Wingdings", "§");

        XSLFTextBox box = slide.createTextBox();
        box.setText("Premier item de la liste");
        XSLFTextParagraph para = box.getTextParagraphs().get(0);
        para.getTextRuns().get(0).setFontFamily("Calibri");
        para.setBulletAutoNumber(AutoNumberingScheme.arabicPeriod, 1);
        // Le paragraphe fraichement numerote hERITE bien Wingdings avant correctif - confirme
        // que ce test reproduit reellement le bug (et pas un cas deja neutre).
        assertEquals("Wingdings", para.getBulletFont(), "precondition : la police heritee doit "
                + "etre Wingdings pour que ce test ait un sens");

        int fixed = AutoNumberBulletFontFixer.fixAutoNumberBulletFonts(slide);

        assertEquals(1, fixed);
        assertEquals("Calibri", para.getBulletFont(), "le numero doit desormais se dessiner dans "
                + "la police du texte, pas dans la police Wingdings heritee du buChar par defaut");
    }

    @Test
    void fixAutoNumberBulletFonts_fallsBackToArialForRunlessParagraph() {
        // Reproduit les 2 paragraphes-espaceurs (invisibles, buAutoNum sans aucun run de
        // texte) du meme slide reel : XSLFTextParagraph#getDefaultFontFamily() replie sur
        // "Arial" quand la liste de runs est vide - meme repli que POI utilise en interne
        // dans DrawTextParagraph#getBullet() quand getBulletFont() renvoie null.
        setMasterOtherStyleLevel1Bullet("Wingdings", "§");

        XSLFTextBox box = slide.createTextBox();
        box.clearText();
        XSLFTextParagraph para = box.addNewTextParagraph();
        para.setBulletAutoNumber(AutoNumberingScheme.arabicPeriod, 8);

        int fixed = AutoNumberBulletFontFixer.fixAutoNumberBulletFonts(slide);

        assertEquals(1, fixed);
        assertEquals("Arial", para.getBulletFont());
    }

    @Test
    void fixAutoNumberBulletFonts_leavesLiteralCharacterBulletUntouched() {
        // Meme heritage Wingdings depuis le masque, mais buChar litteral (pas buAutoNum) :
        // domaine de BulletSymbolFontFixer, pas de celui-ci - getBullet() n'appelle jamais
        // getBulletCharacter() ici de toute facon, voir Javadoc de la classe.
        setMasterOtherStyleLevel1Bullet("Wingdings", "§");

        XSLFTextBox box = slide.createTextBox();
        box.setText("Perimetre");
        XSLFTextParagraph para = box.getTextParagraphs().get(0);
        // Aucun appel a setBulletAutoNumber : la puce heritee est un caractere litteral.

        int fixed = AutoNumberBulletFontFixer.fixAutoNumberBulletFonts(slide);

        assertEquals(0, fixed, "puce a caractere litteral, pas a numerotation automatique");
        assertEquals("Wingdings", para.getBulletFont());
    }

    @Test
    void fixAutoNumberBulletFonts_leavesLocallyDeclaredBulletFontUntouched() {
        // Meme heritage Wingdings depuis le masque, mais CE paragraphe pose explicitement son
        // propre buFont (comme les items "a."/"b." du fichier reel, <a:buFont
        // typeface="+mj-lt"/>) : jamais herite, donc jamais le symptome corrige ici.
        setMasterOtherStyleLevel1Bullet("Wingdings", "§");

        XSLFTextBox box = slide.createTextBox();
        box.setText("Item avec police de puce explicite");
        XSLFTextParagraph para = box.getTextParagraphs().get(0);
        para.setBulletAutoNumber(AutoNumberingScheme.alphaLcPeriod, 1);
        para.setBulletFont("+mj-lt");

        int fixed = AutoNumberBulletFontFixer.fixAutoNumberBulletFonts(slide);

        assertEquals(0, fixed, "buFont deja pose localement pour CE numero precis, rien a corriger");
        assertEquals("+mj-lt", para.getBulletFont());
    }

    @Test
    void fixAutoNumberBulletFonts_leavesNonSymbolInheritedFontUntouched() {
        // La police heritee n'est ni Wingdings ni Symbol (les 2 seules que POI remappe en
        // pictogrammes) : meme heritee "a tort" d'un buChar ancetre, le numero se dessine
        // normalement (aucun symptome visuel), rien a corriger.
        setMasterOtherStyleLevel1Bullet("Arial", "•");

        XSLFTextBox box = slide.createTextBox();
        box.setText("Notification");
        XSLFTextParagraph para = box.getTextParagraphs().get(0);
        para.setBulletAutoNumber(AutoNumberingScheme.arabicPeriod, 1);

        int fixed = AutoNumberBulletFontFixer.fixAutoNumberBulletFonts(slide);

        assertEquals(0, fixed);
        assertEquals("Arial", para.getBulletFont());
    }
}
