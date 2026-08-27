# pptx-slide-to-picture

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
  <artifactId>pptx-slide-to-picture</artifactId>
  <version>0.1.0-SNAPSHOT</version>
</dependency>
```

## Utilisation

```java
import io.github.atlan77c.pptx2picture.OutputFormat;
import io.github.atlan77c.pptx2picture.PptxSlideRenderer;
import io.github.atlan77c.pptx2picture.RenderOptions;

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
- **Écart de métriques de police** : pour certaines polices, Java2D/AWT peut surestimer la hauteur de texte nécessaire par rapport au moteur de rendu natif de PowerPoint (jusqu'à ~30-35% observé dans certains cas — confirmé et quantifié précisément à ~30% sur un fichier réel via un diagnostic dédié, y compris pour une police custom correctement installée et correctement résolue par AWT sous son propre nom), ce qui peut faire déborder une zone de texte hors de sa boîte et chevaucher une forme voisine. **Corrigé par défaut** (`RenderOptions.fixTextOverflow(true)`, activé par défaut) en réduisant la police des formes concernées, différemment selon le mode d'ajustement automatique (`autofit`) configuré sur la forme dans le fichier source : les formes en retrécissement de police (`normAutofit`) ou en "redimensionner la forme selon le texte" (`spAutoFit`) sont systématiquement corrigées (dans les deux cas, PowerPoint ne montre jamais de vrai débordement, donc tout débordement mesuré ici est un artefact du calcul — voir note ci-dessous pour `spAutoFit`) ; les formes sans ajustement (`noAutofit`) ne sont réduites que lorsque le débordement chevaucherait réellement une autre forme de texte, pour rester fidèle aux chevauchements volontaires de l'auteur. Le retrécissement vise une hauteur cible légèrement sous la hauteur réelle de la boîte (marge de sécurité de 3%), pour absorber l'épaisseur du trait de bordure de la forme et un éventuel écart résiduel entre la mesure utilisée pour piloter le correctif et le rendu final — sans cette marge, un cas réel a montré un texte encore visuellement au ras de la bordure malgré une police fortement réduite. Désactivable via `RenderOptions.builder().fixTextOverflow(false)`.

  *Note sur `spAutoFit`* : dans PowerPoint, ce mode fait grandir la boîte plutôt que réduire la police — une première version de ce correctif reproduisait ce comportement en agrandissant la boîte, mais cela s'est révélé provoquer de nouveaux chevauchements avec les formes voisines sur des documents réels (l'agrandissement ne tient pas compte du reste de la mise en page). Réduire la police à la place laisse toutes les autres formes du slide à leur position d'origine : moins fidèle au mécanisme technique de cet autofit, mais plus fidèle au rendu global du diagramme.
- **Titre recouvert par une forme voisine** : le même écart de métriques de police décrit ci-dessus peut, même sans débordement mesuré, faire s'étendre légèrement plus bas le texte d'un titre que chez PowerPoint — suffisant pour le faire chevaucher une image ou une autre forme peinte juste après lui. **Corrigé par défaut** (non désactivable, sans effet sur les slides déjà correctes) en repeignant systématiquement chaque forme de titre par-dessus le reste de la slide une fois le dessin terminé. Un titre est reconnu soit via son type de placeholder OOXML (`title`/`ctrTitle`), soit — quand l'auteur l'a dissocié de son placeholder dans PowerPoint tout en gardant un nom explicite — via le préfixe de son nom, reconnu dans 5 langues : français (« Titre »), anglais (« Title »), espagnol/portugais brésilien (« Título »), italien (« Titolo ») et allemand (« Titel »). Pour reconnaître une langue supplémentaire, ajouter le mot correspondant à la liste `TITLE_NAME_PREFIXES` dans [`TitleRepainter`](src/main/java/io/github/atlan77c/pptx2picture/internal/TitleRepainter.java) (le nom à utiliser se lit dans le volet Sélection de PowerPoint pour une forme de titre dissociée de son placeholder, ou directement dans l'attribut `name` de `<p:nvSpPr><p:cNvPr>` du XML de la forme).
- **Image « découpée selon une forme » rendue intégralement rectangulaire** : une image insérée normalement puis recadrée sur une forme non rectangulaire dans PowerPoint (*Format de l'image > Rogner > Rogner selon la forme* — ellipse, forme libre...) ressortait, avant correctif, entièrement rectangulaire au rendu, recouvrant tout ce qui se trouve derrière elle dans la boîte englobante de son ancre. Cause identifiée dans le code source d'Apache POI lui-même (vérifié sur la version 5.2.5, celle utilisée par ce projet) : `DrawPictureShape` (le dessinateur POI dédié aux images) déclare volontairement n'avoir aucun remplissage propre (`getFillPaint()` renvoie toujours `null`), ce qui saute purement et simplement l'étape de `DrawSimpleShape.draw()` qui calcule et applique le contour géométrique réel de la forme — l'image est alors peinte telle quelle dans le rectangle englobant de son ancre, sans jamais être découpée selon ce contour. **Corrigé par défaut** (non désactivable, sans effet sur les images déjà rectangulaires — insertion normale, ou simple rognage rectangulaire classique `srcRect`) en calculant le contour géométrique réel de la forme et en l'appliquant comme clip avant de peindre l'image — voir [`PictureGeometryClipFixer`](src/main/java/io/github/atlan77c/pptx2picture/internal/PictureGeometryClipFixer.java).
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
