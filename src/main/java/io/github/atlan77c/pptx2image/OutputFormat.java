package io.github.atlan77c.pptx2image;

/**
 * Formats de sortie supportes par {@link PptxSlideRenderer#renderSlideToFile(java.io.File, int, java.io.File, RenderOptions)}
 * (voir {@link RenderOptions.Builder#format(OutputFormat)}).
 */
public enum OutputFormat {

    /**
     * PNG - format raster sans perte, avec canal alpha (transparence).
     * Format par defaut : le meilleur compromis fidelite/compatibilite pour
     * la plupart des usages (affichage web, insertion dans un document).
     */
    PNG,

    /**
     * JPEG - format raster avec compression avec perte, sans transparence
     * (le fond configure via {@link RenderOptions#getBackground()} est
     * toujours rendu opaque dans le fichier produit). Fichiers nettement plus
     * compacts que le PNG, au prix d'artefacts de compression qui deviennent
     * visibles sur les contours nets (texte, formes vectorielles) - a
     * reserver aux cas ou la taille de fichier prime sur la nettete parfaite.
     * Qualite de compression reglable via {@link RenderOptions#getJpegQuality()}.
     */
    JPEG,

    /**
     * SVG - format vectoriel : contrairement au PNG/JPEG, aucun pixel n'est
     * fige a la generation, donc le rendu reste net a n'importe quelle
     * taille/zoom d'affichage. Produit via Apache Batik ({@code SVGGraphics2D}),
     * qui capture les memes appels de dessin que le rendu raster (y compris le
     * correctif de debordement de texte).
     *
     * <p><b>Limite specifique au SVG</b> : le texte est serialise en elements
     * {@code <text>} avec le nom de la police d'origine, pas en glyphes
     * figes comme pour le PNG/JPEG. L'apparence finale depend donc des
     * polices installees sur la machine/le navigateur qui affiche le SVG -
     * si une police du fichier source n'y est pas installee, le rendu visuel
     * (metrique, chasse des caracteres) peut differer de celui obtenu en PNG.
     */
    SVG
}
