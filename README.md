# pptx-slide-to-image

Librairie Java pour convertir une slide d'un fichier PowerPoint (`.pptx`) en image (PNG, JPEG ou SVG), en pur Java via [Apache POI](https://poi.apache.org/) (rendu raster) et [Apache Batik](https://xmlgraphics.apache.org/batik/) (rendu vectoriel SVG) — **sans dépendance à LibreOffice, à PowerPoint (automatisation COM) ni à un navigateur headless**.

## Installation

Le projet n'est pas encore publié sur Maven Central. En attendant, construisez le jar localement (voir [Build](#build)) et installez-le dans votre dépôt Maven local :

```bash
mvn install
```

Puis dans le `pom.xml` de votre projet :

```xml
<dependency>
  <groupId>io.github.Atlan77c</groupId>
  <artifactId>pptx-slide-to-image</artifactId>
  <version>0.1.0-SNAPSHOT</version>
</dependency>
```

## Utilisation

```java
import io.github.atlan77c.pptx2image.OutputFormat;
import io.github.atlan77c.pptx2image.PptxSlideRenderer;
import io.github.atlan77c.pptx2image.RenderOptions;

import java.awt.image.BufferedImage;
import java.io.File;

// Rendu direct vers un fichier PNG (slide n°4, index base 1) - format par défaut
PptxSlideRenderer.renderSlideToFile(new File("presentation.pptx"), 4, new File("slide4.png"));

// Ou récupérer l'image en mémoire pour la manipuler
BufferedImage image = PptxSlideRenderer.renderSlide(new File("presentation.pptx"), 4);

// Avec des options de rendu explicites
RenderOptions options = RenderOptions.builder()
    .scale(2.0f)             // résolution = taille native de la slide x 2 (défaut)
    .fixTextOverflow(true)   // corrige un débordement de texte connu (voir plus bas, défaut: true)
    .build();
PptxSlideRenderer.renderSlideToFile(new File("presentation.pptx"), 4, new File("slide4.png"), options);

// Rendu en JPEG (plus compact que le PNG, sans transparence)
RenderOptions jpegOptions = RenderOptions.builder()
    .format(OutputFormat.JPEG)
    .jpegQuality(0.9f)        // 0.0 (taille mini) à 1.0 (qualité max), défaut 0.92
    .build();
PptxSlideRenderer.renderSlideToFile(new File("presentation.pptx"), 4, new File("slide4.jpg"), jpegOptions);

// Rendu en SVG (vectoriel, net à n'importe quelle taille d'affichage)
RenderOptions svgOptions = RenderOptions.builder().format(OutputFormat.SVG).build();
PptxSlideRenderer.renderSlideToFile(new File("presentation.pptx"), 4, new File("slide4.svg"), svgOptions);
// ou récupérer directement le document SVG en mémoire :
String svgContent = PptxSlideRenderer.renderSlideAsSvg(new File("presentation.pptx"), 4, RenderOptions.defaults());

// Nombre de slides du fichier
int count = PptxSlideRenderer.getSlideCount(new File("presentation.pptx"));
```

## Fidélité du rendu et limites connues

Le texte, les formes vectorielles, les images et les tableaux natifs sont rendus fidèlement. Limites documentées :

- **Les graphiques intégrés (Excel / `XSLFChart`) ne sont pas rendus** par Apache POI — un espace vide apparaît à leur emplacement. C'est une limitation connue d'Apache POI (voir [BUGZILLA-60201](https://bz.apache.org/bugzilla/show_bug.cgi?id=60201)), pas de ce projet. Contournement envisageable (non implémenté ici) : extraire les données du graphique via l'API `XSLFChart`/`XDDFChart` et le redessiner séparément avec une bibliothèque de graphiques Java pure (ex. [JFreeChart](https://www.jfree.org/jfreechart/)).
- **Écart de métriques de police** : pour certaines polices, Java2D/AWT peut surestimer la hauteur de texte nécessaire par rapport au moteur de rendu natif de PowerPoint (jusqu'à ~30-35% observé dans certains cas), ce qui peut faire déborder une zone de texte hors de sa boîte et chevaucher une forme voisine. **Corrigé par défaut** (`RenderOptions.fixTextOverflow(true)`, activé par défaut) : la police des formes concernées est réduite, mais uniquement lorsque le débordement chevaucherait réellement une autre forme de texte — les formes qui débordent "sur le papier" sans risque de collision visuelle ne sont pas touchées, pour rester le plus fidèle possible à la mise en page d'origine. Désactivable via `RenderOptions.builder().fixTextOverflow(false)`.
- **SVG : rendu du texte dépendant des polices installées chez le lecteur** — contrairement au PNG/JPEG (où les glyphes sont figés en pixels à la génération), le SVG produit des éléments `<text>` portant le nom de la police d'origine. Le résultat visuel final (empattement, chasse des caractères) dépend donc des polices disponibles sur la machine ou le navigateur qui affiche le fichier — s'assurer que les polices utilisées dans le `.pptx` source sont bien installées côté lecteur pour un rendu fidèle.

## Build

Nécessite un JDK 17+ et Maven.

```bash
mvn verify
```

Lance la compilation et la suite de tests (JUnit 5), sur des fichiers `.pptx` générés programmatiquement à la volée — aucun fichier `.pptx` externe n'est requis ni embarqué dans le dépôt.

## Comment ça marche

`XSLFSlide` (Apache POI) expose une méthode `draw(Graphics2D)` qui peint récursivement toutes les formes d'une slide. C'est la seule approche testée qui soit 100% Java et sans dépendance système lourde (contrairement aux wrappers LibreOffice/unoconv, à l'automatisation COM PowerPoint, ou aux rendus via navigateur headless type Puppeteer/Chromium). Le correctif de débordement de texte (voir ci-dessus) a été développé et validé itérativement sur des documents réels avant d'être porté dans cette librairie.

Pour le PNG/JPEG, `draw(Graphics2D)` peint dans un `Graphics2D` issu d'un `BufferedImage` (rendu raster classique). Pour le SVG, le même appel `draw(Graphics2D)` reçoit à la place un `SVGGraphics2D` fourni par [Apache Batik](https://xmlgraphics.apache.org/batik/) : cette classe implémente le même contrat `Graphics2D`, mais capture les instructions de dessin pour les sérialiser en SVG au lieu de peindre des pixels. Résultat : le même pipeline de rendu (y compris le correctif de débordement de texte, qui s'appuie lui aussi sur les métriques du `Graphics2D` fourni) produit indifféremment un raster ou du vectoriel, sans code dupliqué. Batik est publié sous licence Apache 2.0 — contrairement à l'alternative la plus connue pour ce besoin (JFreeSVG, sous GPLv3 ou licence commerciale), pleinement compatible avec la licence MIT de ce projet.

## Licence

[MIT](LICENSE)
