package io.github.atlan77c.pptx2image;

/**
 * Erreur levee lorsque la conversion d'une slide pptx en image echoue :
 * fichier illisible/corrompu, index de slide hors bornes, erreur d'ecriture
 * du fichier de sortie, etc. Toujours accompagnee d'un message explicite et,
 * quand la cause est une exception sous-jacente (I/O, parsing OOXML...), de
 * cette cause d'origine (voir {@link #getCause()}).
 */
public class PptxRenderException extends Exception {

    public PptxRenderException(String message) {
        super(message);
    }

    public PptxRenderException(String message, Throwable cause) {
        super(message, cause);
    }
}
