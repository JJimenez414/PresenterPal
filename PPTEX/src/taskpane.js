const OPENAI_API_KEY ="YOUR_OPENAI_API_KEY";

function logDebug(message, data = null) {
   console.log('Debug: ' + message, data || '');
}


Office.onReady(() => {
   logDebug('Office.onReady triggered');
   document.getElementById("executeCommand").onclick = executeCommand;
   document.getElementById("commandInput").addEventListener("keypress", function(event) {
       if (event.key === "Enter") {
           executeCommand();
       }
   });
});


function isImageGenerationCommand(command) {
   const imageKeywords = ['generate image', 'create image', 'make image', 'draw'];
   return imageKeywords.some(keyword => command.toLowerCase().includes(keyword));
}


async function generateImageFromDescription(description) {
   try {
       logDebug('Sending image generation request to OpenAI', { description });


       const response = await fetch('https://api.openai.com/v1/images/generations', {
           method: 'POST',
           headers: {
               'Content-Type': 'application/json',
               'Authorization': 'Bearer ' + OPENAI_API_KEY
           },
           body: JSON.stringify({
               prompt: description,
               n: 1,
               size: '256x256',  // Switch to 256x256 for compatibility
               response_format: 'b64_json'
           })
       });


       if (!response.ok) {
           const errorDetails = await response.text();
           throw new Error('OpenAI API error: ' + response.status + ' ' + errorDetails);
       }


       const data = await response.json();
       logDebug('OpenAI image generation response received', data);


       if (!data.data || !data.data[0]?.b64_json) {
           throw new Error('Invalid response from OpenAI image generation API');
       }


       const base64Image = data.data[0].b64_json.replace(/\s/g, '');
       logDebug('Base64 Image Length: ' + base64Image.length);


       return base64Image;


   } catch (error) {
       logDebug('Error in generateImageFromDescription:', error);
       throw new Error('Failed to generate image with AI: ' + error.message);
   }
}


async function insertImageIntoSlideWithCanvas(base64Image) {
   try {
       // Step 1: Create a canvas and get its context
       const canvas = document.createElement('canvas');
       const ctx = canvas.getContext('2d');


       // Step 2: Create an image element and set the crossOrigin attribute
       const image = new Image();
       image.crossOrigin = 'anonymous'; // Set cross-origin to prevent canvas tainting
       image.src = 'data:image/png;base64,' + base64Image;


       // Wait until the image is loaded
       await new Promise((resolve, reject) => {
           image.onload = () => resolve();
           image.onerror = (err) => reject(new Error('Failed to load image for canvas'));
       });


       // Set canvas size to match the image dimensions
       canvas.width = image.width;
       canvas.height = image.height;


       // Step 3: Draw the image onto the canvas
       ctx.drawImage(image, 0, 0, canvas.width, canvas.height);


       // Step 4: Convert canvas content to a data URL (base64 format)
       const result = canvas.toDataURL('image/png').split(",")[1];


       // Step 5: Use setSelectedDataAsync to insert the image into the slide
       Office.context.document.setSelectedDataAsync(result, {
           coercionType: Office.CoercionType.Image,
           imageLeft: 5,
           imageTop: 5,
           imageWidth: canvas.width / 3,  // Scale down as needed
           imageHeight: canvas.height / 3 // Scale down as needed
       }, function (asyncResult) {
           if (asyncResult.status === Office.AsyncResultStatus.Failed) {
               console.error('Error inserting image:', asyncResult.error.message);
               showStatus("Error inserting image: " + asyncResult.error.message, "error");
           } else {
               console.log('Image inserted successfully');
               showStatus("Image inserted successfully!", "success");
           }
       });
   } catch (error) {
       console.error('Error processing the image for PowerPoint insertion:', error);
       showStatus('Failed to insert image: ' + error.message, "error");
   }
}
async function processWithOpenAI(command, currentState) {
   try {
       logDebug('Sending request to OpenAI', { command, currentState });
      
       const response = await fetch('https://api.openai.com/v1/chat/completions', {
           method: 'POST',
           headers: {
               'Content-Type': 'application/json',
               'Authorization': 'Bearer ' + OPENAI_API_KEY
           },
           body: JSON.stringify({
               model: "gpt-3.5-turbo",
               messages: [
                   {
                       role: "system",
                       content: 'You are a PowerPoint formatting assistant. Convert natural language commands into specific formatting instructions. ' +
                               'Respond only with a JSON object containing the following possible properties: ' +
                               '{ ' +
                               '"color": "hex color code", ' +
                               '"bold": boolean, ' +
                               '"italic": boolean, ' +
                               '"underline": boolean, ' +
                               '"fontSize": number (in points), ' +
                               '"font": "font name", ' +
                               '"error": "error message if command is invalid" ' +
                               '} ' +
                               'Example 1: "make it red and bold" would return {"color": "#FF0000", "bold": true} ' +
                               'Example 2: "remove bold and make it blue" would return {"color": "#0000FF", "bold": false} ' +
                               'Always return valid hex codes for colors.'
                   },
                   {
                       role: "user",
                       content: 'Current state: ' + JSON.stringify(currentState) + '\nCommand: ' + command
                   }
               ],
               temperature: 0.3
           })
       });


       if (!response.ok) {
           throw new Error('OpenAI API error: ' + response.status);
       }


       const data = await response.json();
       logDebug('OpenAI response received', data);


       if (!data.choices || !data.choices[0]?.message?.content) {
           throw new Error('Invalid response from OpenAI');
       }


       const formattingInstructions = JSON.parse(data.choices[0].message.content);
       logDebug('Parsed formatting instructions', formattingInstructions);
       return formattingInstructions;


   } catch (error) {
       logDebug('Error in processWithOpenAI:', error);
       return { error: 'Failed to process command with AI: ' + error.message };
   }
}


async function getCurrentState(shape) {
   try {
       const font = shape.textFrame.textRange.font;
       const state = {
           color: font.color || '#000000',
           bold: font.bold || false,
           italic: font.italic || false,
           underline: font.underline || false,
           fontSize: font.size || 12,
           font: font.name || 'Calibri'
       };
       logDebug('Current state retrieved', state);
       return state;
   } catch (error) {
       logDebug('Error getting current state:', error);
       return {
           color: '#000000',
           bold: false,
           italic: false,
           underline: false,
           fontSize: 12,
           font: 'Calibri'
       };
   }
}


async function applyFormatting(shape, formatting) {
   try {
       logDebug('Applying formatting', formatting);
      
       if (formatting.error) {
           showStatus(formatting.error, "error");
           return;
       }


       const font = shape.textFrame.textRange.font;


       if (formatting.color) {
           font.color = formatting.color;
       }
       if (formatting.bold !== undefined) {
           font.bold = formatting.bold;
       }
       if (formatting.italic !== undefined) {
           font.italic = formatting.italic;
       }
       if (formatting.underline !== undefined) {
           font.underline = formatting.underline;
       }
       if (formatting.fontSize) {
           font.size = formatting.fontSize;
       }
       if (formatting.font) {
           font.name = formatting.font;
       }
      
       logDebug('Formatting applied successfully');
   } catch (error) {
       logDebug('Error applying formatting:', error);
       throw new Error('Failed to apply formatting: ' + error.message);
   }
}


async function executeCommand() {
   logDebug('executeCommand started');
   const commandText = document.getElementById("commandInput").value;
  
   if (!commandText.trim()) {
       showStatus("Please enter a command", "error");
       return;
   }


   showStatus("Processing...", "processing");
  
   try {
       if (isImageGenerationCommand(commandText)) {
           // Extract the actual image description from the command
           const description = commandText.replace(/generate image of|create image of|make image of|draw/gi, '').trim();
           showStatus("Generating image...", "processing");
          
           // Get the base64 image
           const base64Image = await generateImageFromDescription(description);


           // Display in taskpane
           const testImage = document.createElement('img');
           testImage.src = `data:image/png;base64,${base64Image}`;
           testImage.alt = "Generated Image";
           testImage.style.maxWidth = '100%';
           testImage.style.height = 'auto';
           document.getElementById('content-main').appendChild(testImage);


           // Insert into slide using setSelectedDataAsync
           const imageData = `data:image/png;base64,${base64Image}`;
          
           Office.context.document.setSelectedDataAsync(imageData, {
               coercionType: Office.CoercionType.Image,
               imageLeft: 50,
               imageTop: 50,
               imageWidth: 300,
               imageHeight: 300
           }, function (asyncResult) {
               if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                   logDebug('Error inserting image:', asyncResult.error);
                   showStatus("Error inserting image: " + asyncResult.error.message, "error");
               } else {
                   logDebug('Image inserted successfully');
                   showStatus("Image inserted successfully!", "success");
               }
           });


       } else {
           // Your existing text formatting logic
           await PowerPoint.run(async (context) => {
               let slide;
               try {
                   const selectedSlides = context.presentation.getSelectedSlides();
                   selectedSlides.load("items");
                   await context.sync();
                   slide = selectedSlides.items.length > 0 ? selectedSlides.items[0] : context.presentation.slides.getItemAt(0);
               } catch {
                   slide = context.presentation.slides.getItemAt(0);
               }


               let shape;
               try {
                   const selection = slide.shapes.getSelection();
                   selection.load("items");
                   await context.sync();
                   shape = selection.items.length > 0 ? selection.items[0] : slide.shapes.getItemAt(0);
               } catch {
                   shape = slide.shapes.getItemAt(0);
               }


               const currentState = await getCurrentState(shape);
               const formatting = await processWithOpenAI(commandText, currentState);
               await applyFormatting(shape, formatting);
               await context.sync();
               showStatus("Changes applied successfully", "success");
           });
       }
   } catch (error) {
       logDebug('Error in executeCommand:', error);
       showStatus('Error: ' + error.message, "error");
   }


   document.getElementById("commandInput").value = "";
}
function showStatus(message, type) {
   const statusElement = document.getElementById("status");
   statusElement.textContent = message;
   statusElement.className = 'status-message ' + type;
  
   if (type === "success") {
       setTimeout(() => {
           statusElement.className = "status-message";
           statusElement.textContent = "";
       }, 3000);
   }
}