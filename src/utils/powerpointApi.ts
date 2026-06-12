export const insertBlueRectangle = async (): Promise<void> => {
  try {
    await PowerPoint.run(async (context) => {
      const slide = context.presentation.getSelectedSlides().getItemAt(0);

      const rectangle = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle);
      rectangle.left = 50;
      rectangle.top = 50;
      rectangle.width = 250;
      rectangle.height = 100;
      rectangle.fill.setSolidColor("blue");

      rectangle.textFrame.textRange.text = "Template Container";
      rectangle.textFrame.textRange.font.color = "white";
      rectangle.textFrame.textRange.font.bold = true;

      await context.sync();
    });
  } catch (error) {
    console.error("Fehler beim Einfügen des Rechtecks:", error);
  }
};

export const insertGreenCircle = async (): Promise<void> => {
  try {
    await PowerPoint.run(async (context) => {
      const slide = context.presentation.getSelectedSlides().getItemAt(0);

      const circle = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.ellipse);
      circle.left = 100;
      circle.top = 200;
      circle.width = 150;
      circle.height = 150;
      circle.fill.setSolidColor("green");

      circle.textFrame.textRange.text = "Settings";
      circle.textFrame.textRange.font.color = "white";

      await context.sync();
    });
  } catch (error) {
    console.error("Fehler beim Einfügen des Kreises:", error);
  }
};


export const insertSvgIcon = async (svgString: string): Promise<void> => {
  try {
    let accentColor = "#C43E1C"; // Fallback-Farbe (PowerPoint Orange), falls die API fehlschlägt

    // 1. Versuche die Accent Color (Akzent 1) der aktuellen Folie zu laden
    try {
      await PowerPoint.run(async (context) => {
        const slide = context.presentation.getSelectedSlides().getItemAt(0);
        const colorScheme = slide.themeColorScheme;

        // Lade die Akzentfarbe 1 aus dem Master-Theme der Präsentation
        const colorResult = colorScheme.getThemeColor(PowerPoint.ThemeColor.accent1);
        await context.sync();

        accentColor = colorResult.value; // Gibt einen Hex-Code wie "#4472C4" zurück
      });
    } catch (e) {
      console.warn("Konnte Theme Color nicht laden. Nutze Fallback-Farbe.", e);
    }

    // 2. SVG einfärben (ersetzt 'currentColor' durch die geladene Accent Color)
    const coloredSvg = svgString.replace(/currentColor/gi, accentColor);

    // 3. SVG in die Folie einfügen
    // setSelectedDataAsync fügt das Element standardmäßig mittig auf der aktiven Folie ein
    Office.context.document.setSelectedDataAsync(
      coloredSvg,
      { coercionType: Office.CoercionType.XmlSvg },
      (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.error("Fehler beim Einfügen des SVGs:", asyncResult.error.message);
        }
      }
    );
  } catch (error) {
    console.error("Unerwarteter Fehler beim Einfügen des Icons:", error);
  }
};