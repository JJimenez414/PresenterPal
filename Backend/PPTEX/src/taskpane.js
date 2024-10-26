Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    document.getElementById("add-slide-button").onclick = addSlideWithText;
    document.getElementById("move-text-button").onclick = moveTextBox;
    document.getElementById("modify-font-button").onclick = modifyFont;
    document.getElementById("modify-text-button").onclick = modifyText;
    document.getElementById("list-shapes-button").onclick = listShapes;
  }
});

// Function to add a slide with a text box
async function addSlideWithText() {
try {
  await PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items");
    await context.sync();

    const newSlide = slides.add("End");
    newSlide.load("id");
    await context.sync();

    const shapes = newSlide.shapes;
    const textBox = shapes.addTextBox("Hello, PowerPoint!", 100, 100, 400, 50);
    textBox.name = "greetingTextBox"; // Assign a unique name
    textBox.textFrame.textRange.font.size = 24;
    textBox.textFrame.textRange.font.bold = true;

    await context.sync();

    console.log(`Added slide with ID: ${newSlide.id} and text box ID: ${textBox.id}`);
    alert(`Added slide with Text Box ID: ${textBox.id}`);
  });
} catch (error) {
  console.error(error);
  alert("Failed to add slide with text.");
}
}

// Function to move a text box
async function moveTextBox() {
const shapeId = document.getElementById("shape-id").value;
const newLeft = parseFloat(document.getElementById("new-left").value);
const newTop = parseFloat(document.getElementById("new-top").value);

if (!shapeId) {
  alert("Please enter a Shape ID.");
  return;
}

try {
  await PowerPoint.run(async (context) => {
    const slide = context.presentation.slides.getItemAt(0); // Modify as needed
    const shape = slide.shapes.getItem(shapeId);
    
    shape.left = newLeft;
    shape.top = newTop;

    await context.sync();
    console.log(`Moved shape '${shapeId}' to (${newLeft}, ${newTop}).`);
    alert(`Moved shape '${shapeId}' to (${newLeft}, ${newTop}).`);
  });
} catch (error) {
  console.error("Error moving text box:", error);
  alert("Error moving text box. Check the console for details.");
}
}

// Function to modify font properties
async function modifyFont() {
const shapeId = document.getElementById("font-shape-id").value;
const fontSize = parseFloat(document.getElementById("font-size").value);
const fontColor = document.getElementById("font-color").value;
const isBold = document.getElementById("font-bold").checked;

if (!shapeId) {
  alert("Please enter a Shape ID.");
  return;
}

try {
  await PowerPoint.run(async (context) => {
    const slide = context.presentation.slides.getItemAt(0); // Modify as needed
    const shape = slide.shapes.getItem(shapeId);
    
    const textRange = shape.textFrame.textRange;
    textRange.font.size = fontSize;
    textRange.font.color = fontColor;
    textRange.font.bold = isBold;

    await context.sync();
    console.log(`Modified font of shape '${shapeId}'.`);
    alert(`Modified font of shape '${shapeId}'.`);
  });
} catch (error) {
  console.error("Error modifying font:", error);
  alert("Error modifying font. Check the console for details.");
}
}

// Function to modify text content of an existing text box
async function modifyText() {
const shapeId = document.getElementById("text-shape-id").value;
const newText = document.getElementById("new-text").value;

if (!shapeId) {
  alert("Please enter a Shape ID.");
  return;
}

if (!newText) {
  alert("Please enter the new text.");
  return;
}

try {
  await PowerPoint.run(async (context) => {
    const slide = context.presentation.slides.getItemAt(0); // Modify as needed
    const shape = slide.shapes.getItem(shapeId);
    
    // Ensure the shape has a text frame
    if (!shape.textFrame) {
      alert("The specified shape does not contain text.");
      return;
    }

    const textRange = shape.textFrame.textRange;
    textRange.text = newText;

    await context.sync();
    console.log(`Modified text of shape '${shapeId}' to: ${newText}`);
    alert(`Modified text of shape '${shapeId}' to: ${newText}`);
  });
} catch (error) {
  console.error("Error modifying text:", error);
  alert("Error modifying text. Check the console for details.");
}
}

// Utility function to list all shapes on the first slide
async function listShapes() {
try {
  await PowerPoint.run(async (context) => {
    const slide = context.presentation.slides.getItemAt(0);
    const shapes = slide.shapes;
    shapes.load("items/id, items/name, items/type");

    await context.sync();

    if (shapes.items.length === 0) {
      console.log("No shapes found on the slide.");
      alert("No shapes found on the slide.");
      return;
    }

    let shapeInfo = "Shapes on Slide 1:\n";
    shapes.items.forEach((shape, index) => {
      shapeInfo += `\n${index + 1}. ID: ${shape.id}, Name: ${shape.name}, Type: ${shape.type}`;
    });

    console.log(shapeInfo);
    alert(shapeInfo);
  });
} catch (error) {
  console.error("Error listing shapes:", error);
  alert("Error listing shapes. Check the console for details.");
}
}