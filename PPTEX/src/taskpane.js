Office.onReady(() => {
  document.getElementById("changeColorButton").onclick = changeTextColor;
});

function changeTextColor() {
  PowerPoint.run(async (context) => {
    const slide = context.presentation.slides.getItemAt(0); // First slide
    const shape = slide.shapes.getItemAt(0); // First shape
    shape.textFrame.textRange.font.color = "#FF0000"; // Change color to red
    await context.sync();
  });
}
