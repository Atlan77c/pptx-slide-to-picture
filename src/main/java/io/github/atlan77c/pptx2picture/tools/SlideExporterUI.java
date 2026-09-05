package io.github.atlan77c.pptx2picture.tools;

import io.github.atlan77c.pptx2picture.PptxRenderException;
import io.github.atlan77c.pptx2picture.PptxSlideRenderer;

import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;

/**
 * Petite IHM Swing autonome pour exporter en PNG tout ou partie des slides
 * d'un fichier {@code .pptx}, via {@link PptxSlideRenderer} - pensee pour la
 * revue visuelle manuelle (revalider des slides deja corrigees, verifier un
 * nouveau document), pas pour un usage programmatique (voir {@link PptxSlideRenderer}
 * pour l'API de la librairie elle-meme).
 *
 * <p>Ne fait volontairement rien de plus que : choisir un fichier {@code .pptx},
 * choisir tout / une slide / une plage de slides, choisir un dossier de sortie
 * (par defaut le dossier "Telechargements" de l'utilisateur), lancer l'export.
 * Utilise uniquement Swing (present nativement dans le JDK) : aucune dependance
 * supplementaire au projet.
 *
 * <p><b>Lancement</b> : apres {@code mvn install}, executer directement la
 * methode {@link #main} depuis un IDE (Run/Debug), ou en ligne de commande
 * une fois le module compile :
 * <pre>{@code
 * mvn -q compile
 * mvn -q dependency:build-classpath -Dmdep.outputFile=cp.txt
 * java -cp "target/classes;%cp%" io.github.atlan77c.pptx2picture.tools.SlideExporterUI
 * }</pre>
 * (remplacer {@code ;} par {@code :} hors Windows). Aucun plugin Maven
 * supplementaire n'est requis pour que la classe compile et s'execute - un
 * plugin comme {@code exec-maven-plugin} peut etre ajoute au {@code pom.xml}
 * plus tard si un lancement en une seule commande est souhaite, mais n'est pas
 * necessaire pour ce fichier.
 */
public final class SlideExporterUI extends JFrame {

    private static final String DEFAULT_OUTPUT_FOLDER_NAME = "Downloads";

    private final JTextField pptxPathField = new JTextField();
    private final JLabel slideCountLabel = new JLabel(" ");

    private final JRadioButton allSlidesRadio = new JRadioButton("Toutes les slides", true);
    private final JRadioButton singleSlideRadio = new JRadioButton("Une slide precise :");
    private final JRadioButton rangeRadio = new JRadioButton("Une plage :");

    private final JSpinner singleSlideSpinner = new JSpinner(new SpinnerNumberModel(1, 1, 1, 1));
    private final JSpinner rangeFromSpinner = new JSpinner(new SpinnerNumberModel(1, 1, 1, 1));
    private final JSpinner rangeToSpinner = new JSpinner(new SpinnerNumberModel(1, 1, 1, 1));

    private final JTextField outputFolderField = new JTextField(defaultOutputFolder());

    private final JButton launchButton = new JButton("Exporter");
    private final JProgressBar progressBar = new JProgressBar();
    private final JTextArea statusArea = new JTextArea(6, 40);

    /** Nombre de slides du document actuellement charge, ou -1 si aucun/inconnu. */
    private int currentSlideCount = -1;

    public SlideExporterUI() {
        super("Export de slides pptx en images");
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setLayout(new BorderLayout(8, 8));

        add(buildFormPanel(), BorderLayout.NORTH);
        add(buildStatusPanel(), BorderLayout.CENTER);

        pack();
        setMinimumSize(new Dimension(560, getHeight()));
        setLocationRelativeTo(null);

        wireSlideSelectionModeToggle();
        updateLaunchButtonEnabled();
    }

    // ------------------------------------------------------------------
    // Construction de l'IHM
    // ------------------------------------------------------------------

    private JPanel buildFormPanel() {
        JPanel panel = new JPanel(new GridBagLayout());
        panel.setBorder(BorderFactory.createEmptyBorder(10, 10, 0, 10));
        GridBagConstraints c = new GridBagConstraints();
        c.insets = new Insets(4, 4, 4, 4);
        c.fill = GridBagConstraints.HORIZONTAL;
        int row = 0;

        // --- Fichier pptx ---
        c.gridx = 0; c.gridy = row; c.weightx = 0;
        panel.add(new JLabel("Fichier .pptx :"), c);
        c.gridx = 1; c.weightx = 1;
        pptxPathField.setEditable(false);
        panel.add(pptxPathField, c);
        c.gridx = 2; c.weightx = 0;
        JButton browsePptxButton = new JButton("Parcourir...");
        browsePptxButton.addActionListener(this::onBrowsePptx);
        panel.add(browsePptxButton, c);
        row++;

        c.gridx = 1; c.gridy = row; c.gridwidth = 2;
        slideCountLabel.setForeground(Color.DARK_GRAY);
        panel.add(slideCountLabel, c);
        c.gridwidth = 1;
        row++;

        // --- Selection des slides ---
        ButtonGroup modeGroup = new ButtonGroup();
        modeGroup.add(allSlidesRadio);
        modeGroup.add(singleSlideRadio);
        modeGroup.add(rangeRadio);

        c.gridx = 0; c.gridy = row; c.gridwidth = 3;
        panel.add(new JSeparator(), c);
        c.gridwidth = 1;
        row++;

        c.gridx = 0; c.gridy = row; c.gridwidth = 3;
        panel.add(allSlidesRadio, c);
        c.gridwidth = 1;
        row++;

        c.gridx = 0; c.gridy = row;
        panel.add(singleSlideRadio, c);
        c.gridx = 1;
        panel.add(singleSlideSpinner, c);
        row++;

        c.gridx = 0; c.gridy = row;
        panel.add(rangeRadio, c);
        JPanel rangePanel = new JPanel(new FlowLayout(FlowLayout.LEFT, 4, 0));
        rangePanel.add(new JLabel("de"));
        rangePanel.add(rangeFromSpinner);
        rangePanel.add(new JLabel("a"));
        rangePanel.add(rangeToSpinner);
        c.gridx = 1; c.gridwidth = 2;
        panel.add(rangePanel, c);
        c.gridwidth = 1;
        row++;

        // --- Dossier de sortie ---
        c.gridx = 0; c.gridy = row; c.gridwidth = 3;
        panel.add(new JSeparator(), c);
        c.gridwidth = 1;
        row++;

        c.gridx = 0; c.gridy = row; c.weightx = 0;
        panel.add(new JLabel("Dossier de sortie :"), c);
        c.gridx = 1; c.weightx = 1;
        panel.add(outputFolderField, c);
        c.gridx = 2; c.weightx = 0;
        JButton browseOutputButton = new JButton("Parcourir...");
        browseOutputButton.addActionListener(this::onBrowseOutputFolder);
        panel.add(browseOutputButton, c);
        row++;

        // --- Lancement ---
        c.gridx = 0; c.gridy = row; c.gridwidth = 3;
        panel.add(new JSeparator(), c);
        c.gridwidth = 1;
        row++;

        c.gridx = 0; c.gridy = row; c.gridwidth = 3; c.anchor = GridBagConstraints.EAST;
        launchButton.addActionListener(this::onLaunchExport);
        panel.add(launchButton, c);
        c.gridwidth = 1; c.anchor = GridBagConstraints.WEST;
        row++;

        return panel;
    }

    private JPanel buildStatusPanel() {
        JPanel panel = new JPanel(new BorderLayout(4, 4));
        panel.setBorder(BorderFactory.createEmptyBorder(0, 10, 10, 10));
        progressBar.setStringPainted(true);
        panel.add(progressBar, BorderLayout.NORTH);

        statusArea.setEditable(false);
        statusArea.setFont(new Font(Font.MONOSPACED, Font.PLAIN, 12));
        panel.add(new JScrollPane(statusArea), BorderLayout.CENTER);
        return panel;
    }

    private void wireSlideSelectionModeToggle() {
        Runnable updateEnabled = () -> {
            singleSlideSpinner.setEnabled(singleSlideRadio.isSelected());
            rangeFromSpinner.setEnabled(rangeRadio.isSelected());
            rangeToSpinner.setEnabled(rangeRadio.isSelected());
        };
        allSlidesRadio.addActionListener(e -> updateEnabled.run());
        singleSlideRadio.addActionListener(e -> updateEnabled.run());
        rangeRadio.addActionListener(e -> updateEnabled.run());
        updateEnabled.run();
    }

    // ------------------------------------------------------------------
    // Actions utilisateur
    // ------------------------------------------------------------------

    private void onBrowsePptx(ActionEvent event) {
        JFileChooser chooser = new JFileChooser();
        chooser.setFileFilter(new javax.swing.filechooser.FileNameExtensionFilter(
                "Presentations PowerPoint (*.pptx)", "pptx"));
        if (chooser.showOpenDialog(this) != JFileChooser.APPROVE_OPTION) {
            return;
        }
        File selected = chooser.getSelectedFile();
        pptxPathField.setText(selected.getAbsolutePath());
        loadSlideCountInBackground(selected);
    }

    private void onBrowseOutputFolder(ActionEvent event) {
        JFileChooser chooser = new JFileChooser();
        chooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
        File current = new File(outputFolderField.getText().trim());
        if (current.isDirectory()) {
            chooser.setCurrentDirectory(current);
        }
        if (chooser.showOpenDialog(this) != JFileChooser.APPROVE_OPTION) {
            return;
        }
        outputFolderField.setText(chooser.getSelectedFile().getAbsolutePath());
    }

    private void onLaunchExport(ActionEvent event) {
        File pptxFile = new File(pptxPathField.getText().trim());
        if (!pptxFile.isFile()) {
            JOptionPane.showMessageDialog(this, "Choisis d'abord un fichier .pptx valide.",
                    "Fichier manquant", JOptionPane.WARNING_MESSAGE);
            return;
        }

        File outputFolder = new File(outputFolderField.getText().trim());
        if (!outputFolder.exists() && !outputFolder.mkdirs()) {
            JOptionPane.showMessageDialog(this,
                    "Impossible de creer le dossier de sortie :\n" + outputFolder,
                    "Dossier de sortie invalide", JOptionPane.ERROR_MESSAGE);
            return;
        }
        if (!outputFolder.isDirectory()) {
            JOptionPane.showMessageDialog(this,
                    "Le chemin de sortie n'est pas un dossier :\n" + outputFolder,
                    "Dossier de sortie invalide", JOptionPane.ERROR_MESSAGE);
            return;
        }

        List<Integer> slideIndexes = resolveRequestedSlideIndexes();
        if (slideIndexes == null) {
            return; // message d'erreur deja affiche par resolveRequestedSlideIndexes()
        }

        runExportInBackground(pptxFile, outputFolder, slideIndexes);
    }

    /**
     * @return la liste des index de slides (base 1) a exporter selon le mode
     * choisi, ou {@code null} si la selection est invalide (un message
     * d'erreur a deja ete affiche a l'utilisateur dans ce cas).
     */
    private List<Integer> resolveRequestedSlideIndexes() {
        if (currentSlideCount <= 0) {
            JOptionPane.showMessageDialog(this,
                    "Le nombre de slides du document n'a pas pu etre determine - "
                            + "reselectionne le fichier .pptx.",
                    "Document non charge", JOptionPane.WARNING_MESSAGE);
            return null;
        }

        List<Integer> indexes = new ArrayList<>();
        if (allSlidesRadio.isSelected()) {
            for (int i = 1; i <= currentSlideCount; i++) {
                indexes.add(i);
            }
        } else if (singleSlideRadio.isSelected()) {
            indexes.add((Integer) singleSlideSpinner.getValue());
        } else {
            int from = (Integer) rangeFromSpinner.getValue();
            int to = (Integer) rangeToSpinner.getValue();
            if (from > to) {
                JOptionPane.showMessageDialog(this,
                        String.format(Locale.ROOT, "La plage est invalide : %d > %d.", from, to),
                        "Plage invalide", JOptionPane.WARNING_MESSAGE);
                return null;
            }
            for (int i = from; i <= to; i++) {
                indexes.add(i);
            }
        }
        return indexes;
    }

    // ------------------------------------------------------------------
    // Traitement en arriere-plan (SwingWorker : ne bloque jamais l'IHM)
    // ------------------------------------------------------------------

    private void loadSlideCountInBackground(File pptxFile) {
        slideCountLabel.setText("Lecture du document...");
        currentSlideCount = -1;
        updateLaunchButtonEnabled();

        new SwingWorker<Integer, Void>() {
            // Pas de try/catch ici : doInBackground() declare "throws Exception"
            // (contrat de SwingWorker), donc on laisse volontairement remonter
            // n'importe quelle exception - checked (PptxRenderException) ou non
            // (erreur inattendue cote POI/XMLBeans, etc.) - pour ne RIEN avaler
            // silencieusement. Tout est traite de facon uniforme dans done() via
            // get(), qui relance la cause d'origine enveloppee dans une
            // ExecutionException.
            @Override
            protected Integer doInBackground() throws Exception {
                return PptxSlideRenderer.getSlideCount(pptxFile);
            }

            @Override
            protected void done() {
                try {
                    currentSlideCount = get();
                    slideCountLabel.setText(currentSlideCount + " slide(s) detectee(s).");
                    singleSlideSpinner.setModel(new SpinnerNumberModel(1, 1, Math.max(1, currentSlideCount), 1));
                    rangeFromSpinner.setModel(new SpinnerNumberModel(1, 1, Math.max(1, currentSlideCount), 1));
                    rangeToSpinner.setModel(new SpinnerNumberModel(currentSlideCount, 1, Math.max(1, currentSlideCount), 1));
                } catch (Exception e) {
                    currentSlideCount = -1;
                    slideCountLabel.setText(" ");
                    Throwable cause = e.getCause() != null ? e.getCause() : e;
                    JOptionPane.showMessageDialog(SlideExporterUI.this,
                            "Impossible de lire ce fichier :\n" + cause,
                            "Erreur de lecture", JOptionPane.ERROR_MESSAGE);
                    // Trace complete sur la console (utile en lancement IDE/mvn exec:java,
                    // la boite de dialogue ci-dessus ne montre que le message court).
                    cause.printStackTrace();
                }
                updateLaunchButtonEnabled();
            }
        }.execute();
    }

    private void runExportInBackground(File pptxFile, File outputFolder, List<Integer> slideIndexes) {
        setFormEnabled(false);
        progressBar.setMinimum(0);
        progressBar.setMaximum(slideIndexes.size());
        progressBar.setValue(0);
        statusArea.setText("");

        String baseName = stripExtension(pptxFile.getName());
        int digits = String.valueOf(currentSlideCount).length();

        new SwingWorker<Void, String>() {
            private int successCount = 0;
            private int failureCount = 0;

            @Override
            protected Void doInBackground() {
                // 2026-09-05 : le test temporaire de broadenAutofitExemption(true), force
                // pour tous les exports, a ete retire (voir section 26 de
                // conversion_pptx_vers_images.md) - il provoquait des regressions sur des
                // formes NORMAL legitimement rétrécies dans d'autres fichiers/diapositives
                // (placeholders sans forme correspondante au masque, autoformes ordinaires
                // sans autofit declare - voir Javadoc de RenderOptions.Builder#
                // broadenAutofitExemption(boolean)). Retour au comportement par defaut
                // (RenderOptions.defaults(), via la surcharge a 3 arguments) pour tous les
                // fichiers, y compris le slide 16 qui l'avait motive.

                int done = 0;
                for (int slideIndex : slideIndexes) {
                    String fileName = String.format(Locale.ROOT, "%s_slide%0" + digits + "d.png",
                            baseName, slideIndex);
                    File outputFile = new File(outputFolder, fileName);
                    try {
                        // Chaque slide est isolee dans son propre try/catch : un echec
                        // ponctuel (forme non supportee, etc.) n'interrompt pas le lot.
                        PptxSlideRenderer.renderSlideToFile(pptxFile, slideIndex, outputFile);
                        successCount++;
                        publish("OK   slide " + slideIndex + " -> " + fileName);
                    } catch (Exception e) {
                        // Exception large et non plus limitee a PptxRenderException :
                        // une erreur inattendue (pas seulement celles que la librairie
                        // sait anticiper) ne doit pas non plus interrompre le lot ni
                        // rester invisible - voir le meme choix dans loadSlideCountInBackground.
                        failureCount++;
                        publish("ECHEC slide " + slideIndex + " : " + e);
                        e.printStackTrace();
                    }
                    done++;
                    setProgress((int) (100.0 * done / slideIndexes.size()));
                }
                return null;
            }

            @Override
            protected void process(List<String> chunks) {
                for (String line : chunks) {
                    statusArea.append(line + "\n");
                }
                statusArea.setCaretPosition(statusArea.getDocument().getLength());
                progressBar.setValue(progressBar.getValue() + chunks.size());
            }

            @Override
            protected void done() {
                setFormEnabled(true);
                progressBar.setValue(progressBar.getMaximum());
                String summary = String.format(Locale.ROOT,
                        "Termine : %d reussie(s), %d en echec, dans %s",
                        successCount, failureCount, outputFolder.getAbsolutePath());
                statusArea.append("\n" + summary + "\n");
                statusArea.setCaretPosition(statusArea.getDocument().getLength());

                int choice = JOptionPane.showConfirmDialog(SlideExporterUI.this,
                        summary + "\n\nOuvrir le dossier de sortie ?",
                        "Export termine",
                        JOptionPane.YES_NO_OPTION);
                if (choice == JOptionPane.YES_OPTION) {
                    openFolder(outputFolder);
                }
            }
        }.execute();
    }

    // ------------------------------------------------------------------
    // Utilitaires
    // ------------------------------------------------------------------

    private void setFormEnabled(boolean enabled) {
        launchButton.setEnabled(enabled);
        allSlidesRadio.setEnabled(enabled);
        singleSlideRadio.setEnabled(enabled);
        rangeRadio.setEnabled(enabled);
        singleSlideSpinner.setEnabled(enabled && singleSlideRadio.isSelected());
        rangeFromSpinner.setEnabled(enabled && rangeRadio.isSelected());
        rangeToSpinner.setEnabled(enabled && rangeRadio.isSelected());
    }

    private void updateLaunchButtonEnabled() {
        launchButton.setEnabled(currentSlideCount > 0);
    }

    private static void openFolder(File folder) {
        try {
            if (Desktop.isDesktopSupported() && Desktop.getDesktop().isSupported(Desktop.Action.OPEN)) {
                Desktop.getDesktop().open(folder);
            }
        } catch (IOException e) {
            // Non bloquant : le dossier de sortie reste indique dans le resume,
            // l'utilisateur peut toujours y naviguer manuellement.
        }
    }

    private static String stripExtension(String fileName) {
        int dot = fileName.lastIndexOf('.');
        return dot > 0 ? fileName.substring(0, dot) : fileName;
    }

    private static String defaultOutputFolder() {
        File downloads = new File(System.getProperty("user.home"), DEFAULT_OUTPUT_FOLDER_NAME);
        return (downloads.isDirectory() ? downloads : new File(System.getProperty("user.home"))).getAbsolutePath();
    }

    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> new SlideExporterUI().setVisible(true));
    }
}
