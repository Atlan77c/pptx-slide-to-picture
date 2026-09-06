package io.github.atlan77c.pptx2picture.internal;

import org.apache.poi.xslf.usermodel.XSLFGroupShape;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTable;
import org.apache.poi.xslf.usermodel.XSLFTableCell;
import org.apache.poi.xslf.usermodel.XSLFTableRow;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.awt.GraphicsEnvironment;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Set;

/**
 * Filet de securite : substitue en memoire une police declaree mais absente
 * de l'environnement de rendu par un equivalent metriquement compatible
 * reellement installe, avant toute mesure de texte.
 *
 * <p>Constat sur le fichier reel : les presentations traitees declarent des
 * polices Microsoft courantes (Arial, Calibri, Cambria, Times New Roman,
 * Courier New). Sur un poste Windows ou ces polices sont installees, AWT les
 * resout correctement. Sur un serveur Linux de deploiement (conteneur ou VM),
 * ces polices ne sont le plus souvent PAS installees : selon la configuration
 * de fontconfig et du JDK, {@code new Font("Arial", ...)} peut alors soit
 * lever une resolution vers un substitut metriquement proche (ex. Liberation
 * Sans, via les alias de fontconfig), soit retomber silencieusement sur la
 * police logique "Dialog", dont les metriques (ex. DejaVu Sans) different
 * sensiblement de celles d'Arial (~2% observes sur ce projet). Cet ecart,
 * bien que faible en apparence, suffit a deplacer un decoupage de ligne pres
 * du bord d'une forme et a faire manquer les garde-fous de
 * {@link OverflowAwareTextFitter} qui reposent sur des mesures de texte
 * exactes.
 *
 * <p>Mecanisme : pour chaque run de texte dont la police declaree n'est PAS
 * presente dans les polices reellement disponibles de l'environnement de
 * rendu (au sens de {@link GraphicsEnvironment#getAvailableFontFamilyNames()}),
 * on recherche un substitut connu, metriquement compatible, dans une table de
 * correspondance restreinte et documentee. Le substitut n'est applique que
 * s'il est lui-meme disponible. Si la police declaree est deja disponible
 * (cas normal sur un poste Windows avec les polices Microsoft installees),
 * ou si aucun substitut connu n'existe ou n'est disponible, le run n'est pas
 * modifie.
 *
 * <p>Correctif retenu : une simple substitution de {@code fontFamily} en
 * memoire via l'API publique {@link XSLFTextRun#setFontFamily(String)},
 * jamais persistee sur le fichier source. Cette approche est independante du
 * JDK et de la configuration de fontconfig de la machine cible : elle ne
 * depend que de la presence reelle du substitut, verifiee explicitement,
 * plutot que de compter sur un mecanisme de repli de la plateforme qui s'est
 * revele incoherent d'un environnement a l'autre.
 *
 * <p>Limite assumee : la table de correspondance est volontairement
 * restreinte aux polices Microsoft les plus courantes et a leurs equivalents
 * libres les plus repandus (paquet {@code fonts-liberation} et
 * {@code fonts-crosextra-caladea}/{@code carlito}, presents dans la plupart
 * des distributions Linux). Une police non repertoriee et absente de
 * l'environnement de rendu n'est pas substituee : le run conserve sa police
 * declaree et le rendu se degrade alors vers le comportement par defaut de la
 * plateforme, sans regression par rapport a l'existant.
 */
public final class FontSubstitutionFixer {

    private static final Logger LOG = LoggerFactory.getLogger(FontSubstitutionFixer.class);

    /**
     * Table de correspondance police Microsoft -> substitut libre
     * metriquement compatible. Cles en minuscules.
     */
    private static final Map<String, String> KNOWN_SUBSTITUTES = buildSubstituteTable();

    private static Map<String, String> buildSubstituteTable() {
        Map<String, String> m = new HashMap<>();
        m.put("arial", "Liberation Sans");
        m.put("arial narrow", "Liberation Sans Narrow");
        m.put("calibri", "Carlito");
        m.put("cambria", "Caladea");
        m.put("times new roman", "Liberation Serif");
        m.put("courier new", "Liberation Mono");
        return m;
    }

    /**
     * Polices reellement disponibles dans l'environnement de rendu (en
     * minuscules), calculees une seule fois par JVM. Un echec de calcul
     * (environnement graphique degrade) est journalise et traite comme
     * "aucune police disponible", ce qui desactive simplement toute
     * substitution plutot que de faire echouer le rendu.
     */
    private static final Set<String> AVAILABLE_FONT_FAMILIES_LOWER = computeAvailableFontFamilies();

    private static Set<String> computeAvailableFontFamilies() {
        try {
            String[] names = GraphicsEnvironment.getLocalGraphicsEnvironment().getAvailableFontFamilyNames();
            Set<String> lower = new HashSet<>();
            for (String name : names) {
                lower.add(name.toLowerCase(Locale.ROOT));
            }
            return lower;
        } catch (Exception e) {
            LOG.warn("Impossible de lister les polices disponibles ; substitution de police desactivee.", e);
            return new HashSet<>();
        }
    }

    private FontSubstitutionFixer() {
    }

    /**
     * Substitue, sur l'ensemble des formes de texte de la diapositive
     * (y compris dans les groupes et les tableaux), toute police declaree
     * absente de l'environnement de rendu par son equivalent connu si celui-ci
     * est disponible.
     *
     * @param slide diapositive a traiter
     * @return nombre de runs dont la police a ete substituee
     */
    public static int fixMissingFonts(XSLFSlide slide) {
        int fixed = 0;
        for (XSLFShape shape : collectTextShapes(slide.getShapes())) {
            XSLFTextShape ts = (XSLFTextShape) shape;
            for (XSLFTextParagraph paragraph : ts.getTextParagraphs()) {
                for (XSLFTextRun run : paragraph.getTextRuns()) {
                    if (substituteIfMissing(run)) {
                        fixed++;
                    }
                }
            }
        }
        if (fixed > 0 && LOG.isDebugEnabled()) {
            LOG.debug("FontSubstitutionFixer : {} run(s) substitue(s) sur la diapositive '{}'.",
                    fixed, slide.getSlideName());
        }
        return fixed;
    }

    private static boolean substituteIfMissing(XSLFTextRun run) {
        String declared = run.getFontFamily();
        if (declared == null || declared.trim().isEmpty()) {
            return false;
        }
        String declaredLower = declared.toLowerCase(Locale.ROOT);
        if (AVAILABLE_FONT_FAMILIES_LOWER.contains(declaredLower)) {
            // Police reellement installee (ex. Arial sur un poste Windows) : on ne touche a rien.
            return false;
        }
        String substitute = KNOWN_SUBSTITUTES.get(declaredLower);
        if (substitute == null) {
            return false;
        }
        if (!AVAILABLE_FONT_FAMILIES_LOWER.contains(substitute.toLowerCase(Locale.ROOT))) {
            // Le substitut connu n'est lui-meme pas installe sur cette machine : on n'invente rien.
            return false;
        }
        if (LOG.isDebugEnabled()) {
            LOG.debug("Police '{}' absente de l'environnement de rendu -> substitution par '{}'.",
                    declared, substitute);
        }
        run.setFontFamily(substitute);
        return true;
    }

    private static List<XSLFShape> collectTextShapes(List<XSLFShape> shapes) {
        java.util.List<XSLFShape> result = new java.util.ArrayList<>();
        collectTextShapes(shapes, result);
        return result;
    }

    private static void collectTextShapes(List<XSLFShape> shapes, List<XSLFShape> result) {
        for (XSLFShape shape : shapes) {
            if (shape instanceof XSLFGroupShape) {
                collectTextShapes(((XSLFGroupShape) shape).getShapes(), result);
            } else if (shape instanceof XSLFTable) {
                for (XSLFTableRow row : (XSLFTable) shape) {
                    for (XSLFTableCell cell : row) {
                        result.add(cell);
                    }
                }
            } else if (shape instanceof XSLFTextShape) {
                result.add(shape);
            }
        }
    }
}
