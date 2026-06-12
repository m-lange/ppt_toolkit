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



// ... deine bisherigen Funktionen bleiben hier ...

// NEU: 1. Foliengröße auslesen
export const showSlideDimensions = async (): Promise<void> => {
  try {
    await PowerPoint.run(async (context) => {
      // Die PageSetup API ist in neueren PowerPoint-Versionen verfügbar
      if (context.presentation.pageSetup) {
        const pageSetup = context.presentation.pageSetup;
        pageSetup.load("slideWidth, slideHeight");
        await context.sync();

        const width = pageSetup.slideWidth;
        const height = pageSetup.slideHeight;

        // Textbox mit den Maßen einfügen
        const slide = context.presentation.getSelectedSlides().getItemAt(0);
        const textBox = slide.shapes.addTextBox(`Foliengröße:\nBreite: ${width} pt\nHöhe: ${height} pt`);
        textBox.left = 10;
        textBox.top = 10;
        textBox.width = 200;
        textBox.height = 80;
        textBox.fill.setSolidColor("#e1dfdd");

        await context.sync();
      } else {
        console.warn("Die PageSetup API wird in dieser PowerPoint-Version nicht unterstützt.");
      }
    });
  } catch (error) {
    console.error("Fehler beim Auslesen der Foliengröße:", error);
  }
};



export const exportSelectedShapeToJson = async (): Promise<void> => {
  try {
    await PowerPoint.run(async (context) => {
      const shapes = context.presentation.getSelectedShapes();
      shapes.load("items");
      await context.sync();

      if (shapes.items.length === 0) return;

      const shape = shapes.items[0];

      // Lade alle offiziell von Office.js unterstützten Eigenschaften
      shape.load([
        "name", "type", "height", "width", "left", "top",
        "rotation", "lockAspectRatio", // <-- NEU hinzugefügt
        "fill/foregroundColor", "fill/transparency", "fill/type",
        "lineFormat/color", "lineFormat/transparency", "lineFormat/weight", "lineFormat/dashStyle", "lineFormat/visible",
        "textFrame/textRange/text", "textFrame/textRange/font/color", "textFrame/textRange/font/name", "textFrame/textRange/font/size", "textFrame/textRange/font/bold", "textFrame/textRange/font/italic", "textFrame/textRange/font/underline",
        "textFrame/verticalAlignment", "textFrame/leftMargin", "textFrame/rightMargin", "textFrame/topMargin", "textFrame/bottomMargin"
      ]);

      await context.sync();

      // Basis-Eigenschaften, die jedes Shape hat
      const shapeData: any = {
        name: shape.name,
        type: shape.type,
        height: shape.height,
        width: shape.width,
        left: shape.left,
        top: shape.top,
        rotation: shape.rotation,               // <-- NEU
        lockAspectRatio: false  // <-- NEU
      };

      // Try-Catch Blöcke für spezifische Formatierungen (Bilder haben z.B. keinen TextFrame)
      try {
        shapeData.fill = {
          foregroundColor: shape.fill.foregroundColor,
          transparency: shape.fill.transparency,
          type: shape.fill.type
        };
      } catch(e) {}

      try {
        shapeData.lineFormat = {
          color: shape.lineFormat.color,
          transparency: shape.lineFormat.transparency,
          weight: shape.lineFormat.weight,
          dashStyle: shape.lineFormat.dashStyle,
          visible: shape.lineFormat.visible
        };
      } catch(e) {}

      try {
        shapeData.textFrame = {
          verticalAlignment: shape.textFrame.verticalAlignment,
          leftMargin: shape.textFrame.leftMargin,
          rightMargin: shape.textFrame.rightMargin,
          topMargin: shape.textFrame.topMargin,
          bottomMargin: shape.textFrame.bottomMargin,
          text: shape.textFrame.textRange.text,
        };
        try {
          shapeData.textFrame.font = {
            color: shape.textFrame.textRange.font.color,
            name: shape.textFrame.textRange.font.name,
            size: shape.textFrame.textRange.font.size,
            bold: shape.textFrame.textRange.font.bold,
            italic: shape.textFrame.textRange.font.italic,
            underline: shape.textFrame.textRange.font.underline
          };
        } catch(e) {}
      } catch(e) {}

      const jsonString = JSON.stringify(shapeData, null, 2);

      // Textbox mit dem JSON formatieren und einfügen
      const slide = context.presentation.getSelectedSlides().getItemAt(0);
      const textBox = slide.shapes.addTextBox(jsonString);

      textBox.left = shape.left + 20;
      textBox.top = shape.top + 20;
      textBox.width = 350;

      // Styling der JSON-Textbox (Code-Look)
      textBox.textFrame.textRange.font.size = 10;
      textBox.textFrame.textRange.font.name = "Courier New";
      textBox.fill.setSolidColor("#F3F2F1");
      textBox.lineFormat.color = "#8A8886";
      textBox.lineFormat.weight = 1;

      await context.sync();
    });
  } catch (error) {
    console.error("Fehler beim Exportieren des Shapes:", error);
  }
};