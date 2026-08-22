# pptx-slide-to-image

Librairie Java pour convertir une slide d'un fichier PowerPoint (`.pptx`) en image (PNG), en pur Java via [Apache POI](https://poi.apache.org/) — **sans dépendance à LibreOffice, à PowerPoint (automatisation COM) ni à un navigateur headless**.

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
import io.github.atlan77c.pptx2image.PptxSlideRenderer;
import io.github.atlan77c.pptx2image.RenderOptions;

import java.awt.image.BufferedImage;
import java.io.File;

// Rendu direct vers un fichier PNG (slide n°4, index base 1)
PptxSlideRenderer.renderSlideToFile(new File("presentation.pptx"), 4, new File("slide4.png"));

// Ou récupérer l'image en mémoire pour la manipuler
BufferedImage image = PptxSlideRenderer.renderSlide(new File("presentation.pptx"), 4);

// Avec des options de rendu explicites
RenderOptions options = RenderOptions.builder()
    .scale(2.0f)             // résolution = taille native de la slide x 2 (défaut)
    .fixTextOverflow(true)   // corrige un débordement de texte connu (voir plus bas, défaut: true)
    .build();
PptxSlideRenderer.renderSlideToFile(new File("presentation.pptx"), 4, new File("slide4.png"), options);

// Nombre de slides du fichier
int count = PptxSlideRenderer.getSlideCount(new File("presentation.pptx"));
```

## Fidélité du rendu et limites connues

Le texte, les formes vectorielles, les images et les tableaux natifs sont rendus fidèlement. Deux limites sont documentées :

- **Les graphiques intégrés (Excel / `XSLFChart`) ne sont pas rendus** par Apache POI — un espace vide apparaît à leur emplacement. C'est une limitation connue d'Apache POI (voir [BUGZILLA-60201](https://bz.apache.org/bugzilla/show_bug.cgi?id=60201)), pas de ce projet. Contournement envisageable (non implémenté ici) : extraire les données du graphique via l'API `XSLFChart`/`XDDFChart` et le redessiner séparément avec une bibliothèque de graphiques Java pure (ex. [JFreeChart](https://www.jfree.org/jfreechart/)).
- **Écart de métriques de police** : pour certaines polices, Java2D/AWT peut surestimer la hauteur de texte nécessaire par rapport au moteur de rendu natif de PowerPoint (observé jusqu'à ~30-35% avec la police "PoliceX"), ce qui peut faire déborder une zone de texte hors de sa boîte et chevaucher une forme voisine. **Corrigé par défaut** (`RenderOptions.fixTextOverflow(true)`, activé par défaut) : la police des formes concernées est réduite, mais uniquement lorsque le débordement chevaucherait réellement une autre forme de texte — les formes qui débordent "sur le papier" sans risque de collision visuelle ne sont pas touchées, pour rester le plus fidèle possible à la mise en page d'origine. Désactivable via `RenderOptions.builder().fixTextOverflow(false)`.

## Build

Nécessite un JDK 17+ et Maven.

```bash
mvn verify
```

Lance la compilation et la suite de tests (JUnit 5), sur des fichiers `.pptx` générés programmatiquement à la volée — aucun fichier `.pptx` externe n'est requis ni embarqué dans le dépôt.

## Comment ça marche

`XSLFSlide` (Apache POI) expose une méthode `draw(Graphics2D)` qui peint récursivement toutes les formes d'une slide dans une image. C'est la seule approche testée qui soit 100% Java et sans dépendance système lourde (contrairement aux wrappers LibreOffice/unoconv, à l'automatisation COM PowerPoint, ou aux rendus via navigateur headless type Puppeteer/Chromium). Le correctif de débordement de texte (voir ci-dessus) a été développé et validé itérativement sur des documents réels avant d'être porté dans cette librairie.

## Licence

[MIT](LICENSE)
