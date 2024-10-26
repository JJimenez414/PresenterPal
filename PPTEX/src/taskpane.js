Office.onReady(() => {
  document.getElementById("executeCommand").onclick = executeCommand;
  document.getElementById("commandInput").addEventListener("keypress", function(event) {
      if (event.key === "Enter") {
          executeCommand();
      }
  });
});

async function executeCommand() {
  const commandText = document.getElementById("commandInput").value.toLowerCase();
  const statusElement = document.getElementById("status");
  
  try {
      await PowerPoint.run(async (context) => {
          // Get the selected slide, if none selected, get the first slide
          let slide;
          try {
              slide = context.presentation.getSelectedSlides();
              await context.sync();
              if (slide.items.length === 0) {
                  slide = context.presentation.slides.getItemAt(0);
              } else {
                  slide = slide.items[0];
              }
          } catch {
              slide = context.presentation.slides.getItemAt(0);
          }
          
          // Get selected shape or first shape if none selected
          let shape;
          try {
              const selection = slide.shapes.getSelection();
              await context.sync();
              if (selection.items.length === 0) {
                  shape = slide.shapes.getItemAt(0);
              } else {
                  shape = selection.items[0];
              }
          } catch {
              shape = slide.shapes.getItemAt(0);
          }

          // Process commands
          if (commandText.includes("red")) {
              shape.textFrame.textRange.font.color = "#FF0000";
              showStatus("Text color changed to red", "success");
          } 
          else if (commandText.includes("blue")) {
              shape.textFrame.textRange.font.color = "#0000FF";
              showStatus("Text color changed to blue", "success");
          }
          else if (commandText.includes("green")) {
              shape.textFrame.textRange.font.color = "#00FF00";
              showStatus("Text color changed to green", "success");
          }
          else if (commandText.includes("black")) {
              shape.textFrame.textRange.font.color = "#000000";
              showStatus("Text color changed to black", "success");
          }
          else if (commandText.includes("bold")) {
              shape.textFrame.textRange.font.bold = true;
              showStatus("Text changed to bold", "success");
          }
          else if (commandText.includes("unbold") || commandText.includes("remove bold")) {
              shape.textFrame.textRange.font.bold = false;
              showStatus("Bold removed from text", "success");
          }
          else if (commandText.includes("italic")) {
              shape.textFrame.textRange.font.italic = true;
              showStatus("Text changed to italic", "success");
          }
          else if (commandText.includes("underline")) {
              shape.textFrame.textRange.font.underline = true;
              showStatus("Text underlined", "success");
          }
          else {
              showStatus("Command not recognized. Try commands like 'make it red' or 'make it bold'", "error");
              return;
          }

          await context.sync();
      });
  } catch (error) {
      showStatus("Error: " + error.message, "error");
  }

  // Clear input after executing command
  document.getElementById("commandInput").value = "";
}

function showStatus(message, type) {
  const statusElement = document.getElementById("status");
  statusElement.textContent = message;
  statusElement.className = "status-message " + type;
  
  // Hide status message after 3 seconds
  setTimeout(() => {
      statusElement.className = "status-message";
  }, 3000);
}