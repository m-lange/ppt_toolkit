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