package io.github.atlan77c.pptx2picture.internal;

import org.apache.poi.common.usermodel.fonts.FontGroup;
import org.apache.poi.sl.usermodel.PaintStyle;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;

import java.util.ArrayList;
import java.util.List;

/**
 * Insere un saut de ligne FORCE ({@link XSLFTextParagraph#addLineBreak()})
 * exactement au point de depassement horizontal detecte par {@link
 * org.apache.poi.sl.draw.HorizontalOverflowLineBreakFixer} (le ou les
 * derniers mots d'une ligne trop large), plutot que de retrecir la police -
 * preserve la taille de police d'origine, donc plus fidele visuellement a
 * PowerPoint que l'approche "retrecissement" pour ce point precis. Cas reels
 * confirmes ayant motive ce correctif (meme document interne reel que dans le reste
 * de ce paquetage) :
 * slide 57 ("...trop perçus  et la rupture...", reflow horizontal revele par le
 * retrecissement vertical existant - voir {@link OverflowAwareTextFitter} -
 * qui recalcule un decoupage de ligne different) et slide 68 (Rectangle 9/10/12).
 *
 * <p><b>Contrainte API respectee</b> (meme technique que {@link
 * SymbolFontRunFixer}, section "Limite assumee" de sa Javadoc) : {@code
 * XSLFTextParagraph#addNewTextRun()}/{@code addLineBreak()} ne peuvent
 * qu'AJOUTER en fin de paragraphe - une manipulation directe de l'arbre XML
 * sous-jacent ({@code CTTextParagraph.insertNewR()}) desynchronise l'etat
 * interne de POI (deja constate et documente dans ce projet, voir Javadoc de
 * {@link SymbolFontRunFixer}). Donc, pour inserer "au milieu" : (1) le run
 * contenant le point de coupure est TRONQUE en place (mutation de contenu, pas
 * de structure) a son prefixe ; (2) tous les runs ORIGINAUX suivants (y
 * compris le suffixe du run tronque) sont VIDES ({@code setText("")}) ; (3) le
 * contenu retire est RE-AJOUTE dans l'ordre en fin de paragraphe via l'API
 * publique, avec un {@code addLineBreak()} insere juste avant le premier
 * segment pousse a la ligne suivante - chaque nouveau run recoit
 * EXPLICITEMENT la taille/gras/italique/police latine/couleur du run ORIGINAL
 * dont il reprend le texte (un nouveau run ne les herite pas par defaut, meme
 * lecon deja tiree pour {@link SymbolFontRunFixer}).
 *
 * <p><b>{@code XSLFLineBreak} est package-private</b> ({@code
 * org.apache.poi.xslf.usermodel}) : impossible de le referencer par son nom
 * depuis ce paquetage {@code internal}, ni via {@code instanceof} - detecte
 * ici par nom de classe ({@link #isLineBreak}). Point important : un saut de
 * ligne DEJA present dans la zone reconstruite (ex. insere par une coupure
 * precedente sur le meme paragraphe, quand plusieurs sauts de ligne sont
 * necessaires) ne doit jamais recevoir {@code setText("")} - {@code
 * XSLFLineBreak#setText()} leve toujours {@code IllegalStateException} (texte
 * fixe a {@code "\n"}) - il est prealablement identifie et REJOUE via un
 * nouveau {@code addLineBreak()} (bug rencontre et corrige sur la slide 68,
 * Rectangle 10, ou plusieurs sauts de ligne successifs etaient necessaires
 * dans le meme paragraphe).
 */
public final class ForcedLineBreakFixer {

    private ForcedLineBreakFixer() {
    }

    /** Point de coupure localise dans le texte concatene des runs d'un paragraphe. */
    public static final class SplitPoint {
        public final int splitRunIdx;
        public final int localOffsetInRun;

        SplitPoint(int splitRunIdx, int localOffsetInRun) {
            this.splitRunIdx = splitRunIdx;
            this.localOffsetInRun = localOffsetInRun;
        }
    }

    /**
     * Localise (sans rien modifier) le point de coupure correspondant aux
     * {@code wordsToDrop} derniers mots de {@code overflowingLineText} (une
     * ligne de {@code para} telle que decoupee par {@code breakText()}) -
     * permet un essai non destructif (voir {@link #previewTruncatedWidth}) ou
     * une insertion directe (voir {@link #applySplit}).
     *
     * @return le point de coupure, ou {@code null} si introuvable (pas assez
     *         de mots, ligne non retrouvee dans le texte du paragraphe, ou
     *         coupure tombant sur un saut de ligne deja present).
     */
    public static SplitPoint locateSplitPoint(XSLFTextParagraph para, String overflowingLineText, int wordsToDrop) {
        List<XSLFTextRun> originalRuns = para.getTextRuns();
        StringBuilder fullBuilder = new StringBuilder();
        int[] runStartOffset = new int[originalRuns.size()];
        for (int i = 0; i < originalRuns.size(); i++) {
            runStartOffset[i] = fullBuilder.length();
            String t = originalRuns.get(i).getRawText();
            fullBuilder.append(t == null ? "" : t);
        }
        String full = fullBuilder.toString();

        int lineStart = full.indexOf(overflowingLineText);
        if (lineStart < 0) {
            return null;
        }

        int localSplit = findSplitOffset(overflowingLineText, wordsToDrop);
        if (localSplit <= 0) {
            return null;
        }
        int absoluteSplit = lineStart + localSplit;

        int splitRunIdx = -1;
        for (int i = 0; i < originalRuns.size(); i++) {
            int start = runStartOffset[i];
            int end = start + lengthOf(originalRuns.get(i));
            if (absoluteSplit >= start && absoluteSplit < end) {
                splitRunIdx = i;
                break;
            }
            if (absoluteSplit == end && i == originalRuns.size() - 1) {
                splitRunIdx = i;
            }
        }
        if (splitRunIdx < 0) {
            return null;
        }
        if (isLineBreak(originalRuns.get(splitRunIdx))) {
            // Coupure tombant sur un saut de ligne deja present (ex. insere par
            // une coupure precedente sur ce meme paragraphe) - pas un point de
            // coupure exploitable, voir Javadoc de applySplit().
            return null;
        }
        int localOffsetInRun = absoluteSplit - runStartOffset[splitRunIdx];
        return new SplitPoint(splitRunIdx, localOffsetInRun);
    }

    /**
     * Essai NON DESTRUCTIF : tronque temporairement (en place, {@code
     * setText}) le run de coupure a son prefixe ET VIDE temporairement tous
     * les runs suivants (comme le fera {@link #applySplit} pour de vrai) -
     * ainsi le texte d'essai du paragraphe s'arrete EXACTEMENT au point de
     * coupure candidat, ce qui rend {@code measure} (typiquement un nouveau
     * {@code breakText()} + lecture de la DERNIERE ligne non vide, forcement
     * sans ambiguite puisque rien ne la suit) fiable MEME quand la ligne en
     * depassement n'est pas la toute premiere ligne du paragraphe - bug
     * corrige le 2026-09-06 (slide 68, Rectangle 10) : verifier seulement la
     * ligne 0 revenait a valider n'importe quel {@code wordsToDrop} des qu'un
     * depassement PLUS LOIN dans le paragraphe etait vise (paragraphe
     * necessitant plusieurs sauts de ligne successifs), produisant des
     * coupures en cascade mal placees (le mot "sur" isole seul sur sa ligne).
     * Restaure systematiquement tous les textes d'origine (y compris si
     * {@code measure} leve une exception) - jamais de mutation durable avant
     * d'avoir trouve le bon candidat.
     */
    public static double previewTruncatedWidth(XSLFTextParagraph para, SplitPoint sp, java.util.function.Supplier<Double> measure) {
        List<XSLFTextRun> runs = para.getTextRuns();
        XSLFTextRun splitRun = runs.get(sp.splitRunIdx);
        String originalText = splitRun.getRawText() == null ? "" : splitRun.getRawText();
        List<XSLFTextRun> laterRuns = new ArrayList<>();
        List<String> laterOriginalTexts = new ArrayList<>();
        for (int i = sp.splitRunIdx + 1; i < runs.size(); i++) {
            XSLFTextRun r = runs.get(i);
            if (isLineBreak(r)) {
                continue;
            }
            laterRuns.add(r);
            laterOriginalTexts.add(r.getRawText());
        }
        try {
            splitRun.setText(sp.localOffsetInRun < originalText.length()
                    ? originalText.substring(0, sp.localOffsetInRun) : originalText);
            for (XSLFTextRun r : laterRuns) {
                r.setText("");
            }
            return measure.get();
        } finally {
            splitRun.setText(originalText);
            for (int i = 0; i < laterRuns.size(); i++) {
                laterRuns.get(i).setText(laterOriginalTexts.get(i));
            }
        }
    }

    /**
     * Insere un saut de ligne force juste avant les {@code wordsToDrop}
     * derniers mots de {@code overflowingLineText} (une ligne de {@code para}
     * telle que decoupee par {@code breakText()}), en localisant le point de
     * coupure dans le texte concatene des runs du paragraphe.
     *
     * @return {@code true} si l'insertion a reussi (point de coupure
     *         localise et paragraphe reconstruit), {@code false} sinon (rien
     *         modifie).
     */
    public static boolean insertBreakBeforeLastWords(XSLFTextParagraph para, String overflowingLineText, int wordsToDrop) {
        SplitPoint sp = locateSplitPoint(para, overflowingLineText, wordsToDrop);
        if (sp == null) {
            return false;
        }
        return applySplit(para, sp);
    }

    /**
     * Realise la coupure DEFINITIVE (tronque + reconstruit la fin du
     * paragraphe avec un saut de ligne force) a un point de coupure deja
     * localise - voir Javadoc de la classe pour la technique (jamais {@code
     * insertNewR()}, toujours l'API publique en mode "vider + rajouter en
     * fin").
     */
    public static boolean applySplit(XSLFTextParagraph para, SplitPoint sp) {
        List<XSLFTextRun> originalRuns = new ArrayList<>(para.getTextRuns());
        int splitRunIdx = sp.splitRunIdx;
        int localOffsetInRun = sp.localOffsetInRun;

        XSLFTextRun splitRun = originalRuns.get(splitRunIdx);
        if (isLineBreak(splitRun)) {
            // Coupure tombant exactement sur un saut de ligne DEJA existant
            // (ex. une coupure precedente sur ce meme paragraphe, voir Javadoc
            // de la classe) : rien de pertinent a scinder ici.
            return false;
        }
        String splitRunText = splitRun.getRawText() == null ? "" : splitRun.getRawText();

        // Capture des segments a re-ajouter APRES le point de coupure, dans l'ordre,
        // AVANT toute mutation (formatage individuel de chaque run d'origine). Un
        // saut de ligne DEJA present dans cette zone (ex. insere par une coupure
        // precedente sur ce meme paragraphe, dans un autre passage de la boucle
        // appelante) est un marqueur special (texte=null) : on le REJOUE via
        // addLineBreak(), jamais via addNewTextRun()/setText() (XSLFLineBreak
        // refuse tout texte autre que "\n" - voir Javadoc de la classe).
        List<String> pendingText = new ArrayList<>();
        List<Double> pendingSize = new ArrayList<>();
        List<Boolean> pendingBold = new ArrayList<>();
        List<Boolean> pendingItalic = new ArrayList<>();
        List<String> pendingFamily = new ArrayList<>();
        List<PaintStyle> pendingColor = new ArrayList<>();

        String suffixOfSplitRun = localOffsetInRun < splitRunText.length() ? splitRunText.substring(localOffsetInRun) : "";
        if (!suffixOfSplitRun.isEmpty()) {
            pendingText.add(suffixOfSplitRun);
            pendingSize.add(splitRun.getFontSize());
            pendingBold.add(splitRun.isBold());
            pendingItalic.add(splitRun.isItalic());
            pendingFamily.add(splitRun.getFontFamily(FontGroup.LATIN));
            pendingColor.add(splitRun.getFontColor());
        }
        for (int i = splitRunIdx + 1; i < originalRuns.size(); i++) {
            XSLFTextRun r = originalRuns.get(i);
            if (isLineBreak(r)) {
                pendingText.add(null); // marqueur "saut de ligne" - voir Javadoc ci-dessus.
                pendingSize.add(null);
                pendingBold.add(null);
                pendingItalic.add(null);
                pendingFamily.add(null);
                pendingColor.add(null);
                continue;
            }
            String t = r.getRawText();
            if (t == null || t.isEmpty()) {
                continue;
            }
            pendingText.add(t);
            pendingSize.add(r.getFontSize());
            pendingBold.add(r.isBold());
            pendingItalic.add(r.isItalic());
            pendingFamily.add(r.getFontFamily(FontGroup.LATIN));
            pendingColor.add(r.getFontColor());
        }
        if (pendingText.isEmpty()) {
            // Rien a pousser a la ligne suivante : coupure inutile.
            return false;
        }

        // Mutation : tronque le run de coupure, vide les runs originaux suivants
        // (les sauts de ligne deja presents dans cette zone sont laisses tels
        // quels - XSLFLineBreak n'a de toute facon aucun contenu textuel propre).
        splitRun.setText(localOffsetInRun < splitRunText.length() ? splitRunText.substring(0, localOffsetInRun) : splitRunText);
        for (int i = splitRunIdx + 1; i < originalRuns.size(); i++) {
            XSLFTextRun r = originalRuns.get(i);
            if (isLineBreak(r)) {
                continue;
            }
            r.setText("");
        }

        // Re-ajoute en fin de paragraphe, via l'API publique uniquement (voir Javadoc).
        para.addLineBreak();
        for (int i = 0; i < pendingText.size(); i++) {
            if (pendingText.get(i) == null) {
                para.addLineBreak();
                continue;
            }
            XSLFTextRun newRun = para.addNewTextRun();
            newRun.setText(pendingText.get(i));
            if (pendingSize.get(i) != null) {
                newRun.setFontSize(pendingSize.get(i));
            }
            if (pendingBold.get(i) != null) {
                newRun.setBold(pendingBold.get(i));
            }
            if (pendingItalic.get(i) != null) {
                newRun.setItalic(pendingItalic.get(i));
            }
            if (pendingFamily.get(i) != null) {
                newRun.setFontFamily(pendingFamily.get(i), FontGroup.LATIN);
            }
            if (pendingColor.get(i) != null) {
                newRun.setFontColor(pendingColor.get(i));
            }
        }
        return true;
    }

    /**
     * Longueur EN CARACTERES du texte brut de {@code r}, telle qu'ajoutee au
     * texte concatene du paragraphe dans {@link #locateSplitPoint} (doit
     * rester coherente avec {@code fullBuilder.append(r.getRawText())} pour
     * que les offsets calcules restent alignes - donc PAS de cas particulier
     * pour un saut de ligne ici, contrairement a {@link #applySplit}).
     */
    private static int lengthOf(XSLFTextRun r) {
        String t = r.getRawText();
        return t == null ? 0 : t.length();
    }

    /**
     * {@code XSLFLineBreak} (paquetage {@code org.apache.poi.xslf.usermodel})
     * est package-private : impossible de le referencer par son nom depuis ce
     * paquetage {@code internal}, d'ou cette detection par nom de classe -
     * seule methode disponible pour distinguer un run "saut de ligne" d'un
     * {@code CTRegularTextRun} ordinaire depuis l'exterieur de son paquetage.
     */
    private static boolean isLineBreak(XSLFTextRun r) {
        return "XSLFLineBreak".equals(r.getClass().getSimpleName());
    }

    /**
     * @return l'offset LOCAL (dans {@code lineText}) juste avant le
     *         {@code wordsToDrop}-ieme mot en partant de la fin, ou -1 si pas
     *         assez de mots. Un "mot" est une sequence maximale de caracteres
     *         non-espace ; l'espace precedant le mot coupe est inclus dans la
     *         partie retiree (n'est pas rendu apres coupure - sans
     *         consequence visuelle, l'espace en fin de ligne n'etant jamais
     *         visible).
     */
    private static int findSplitOffset(String lineText, int wordsToDrop) {
        int n = lineText.length();
        int i = n;
        int wordsFound = 0;
        while (i > 0 && wordsFound < wordsToDrop) {
            // Saute les espaces en fin.
            while (i > 0 && Character.isWhitespace(lineText.charAt(i - 1))) {
                i--;
            }
            int wordEnd = i;
            while (i > 0 && !Character.isWhitespace(lineText.charAt(i - 1))) {
                i--;
            }
            int wordStart = i;
            if (wordStart == wordEnd) {
                break; // aucun mot trouve
            }
            wordsFound++;
        }
        if (wordsFound < wordsToDrop) {
            return -1;
        }
        // i pointe maintenant juste apres le dernier espace precedant le premier
        // mot retire - c'est le point de coupure (espace(s) de separation exclus
        // de la partie conservee).
        while (i > 0 && Character.isWhitespace(lineText.charAt(i - 1))) {
            i--;
        }
        return i;
    }
}
