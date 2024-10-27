// OpenAI API Key
const OPENAI_API_KEY = "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";

// Debug logging helper
function logDebug(message, data = null) {
    console.log(`Debug: ${message}`, data || '');
}

// Wait for Office to be ready
Office.onReady(() => {
    logDebug('Office.onReady triggered');


    // Initialize ChatApp if available
    if (window.ChatApp) {
        window.ChatApp.initialize();
    } else {
        console.error("ChatApp not found");
    }

    // Set up event listeners once Office is ready
    setupEventListeners();
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
        const canvas = document.createElement('canvas');
        const ctx = canvas.getContext('2d');
 
 
        const image = new Image();
        image.crossOrigin = 'anonymous'; // Set cross-origin to prevent canvas tainting
        image.src = 'data:image/png;base64,' + base64Image;
 
 
        await new Promise((resolve, reject) => {
            image.onload = () => resolve();
            image.onerror = (err) => reject(new Error('Failed to load image for canvas'));
        });
 
 
        canvas.width = image.width;
        canvas.height = image.height;
 
        ctx.drawImage(image, 0, 0, canvas.width, canvas.height);
 
        const result = canvas.toDataURL('image/png').split(",")[1];
 
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

function setupEventListeners() {
    const commandInput = document.getElementById("commandInput");
    const executeButton = document.getElementById("executeCommand");
    
    if (commandInput) {
        commandInput.addEventListener("keypress", function(event) {
            if (event.key === "Enter") {
                event.preventDefault();
                executeCommand();
            }
        });
        logDebug('Command input event listener attached');
    } else {
        logDebug('Command input element not found');
    }

    if (executeButton) {
        executeButton.addEventListener("click", function(event) {
            event.preventDefault();
            executeCommand();
        });
        logDebug('Execute button event listener attached');
    } else {
        logDebug('Execute button element not found');
    }
}

async function processWithOpenAI(command, currentState) {
    try {
        logDebug('Sending request to OpenAI', { command, currentState });
        
        const response = await fetch('https://api.openai.com/v1/chat/completions', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${OPENAI_API_KEY}`
            },
            body: JSON.stringify({
                model: "gpt-3.5-turbo",
                messages: [
                    {
                        role: "system",
                        content: `You are a PowerPoint formatting and creation assistant. Convert natural language commands into specific formatting instructions or creation commands. 
                        When multiple actions are requested, return an array of instructions.
                        Respond with a JSON object containing either a single instruction or an array of instructions.
                        Each instruction can have these properties:
                    
                        For creating new text:
                        {
                            "action": "create",
                            "type": "textbox",
                            "text": "text content",
                            "bulletPoints": ["point 1", "point 2", "etc"], // Array of bullet points if requested
                            "hasBullets": boolean // true if bullet points are requested
                        }
                    
                        For moving text:
                        {
                            "position": "center" | "left" | "right" | "up" | "down",
        "distance": number (optional, default 20 for up/down movements
                        }
                    
                        For formatting existing text:
                        {
                            "action": "format",
                            "color": "hex color code",
                            "bold": boolean,
                            "italic": boolean,
                            "underline": boolean,
                            "fontSize": number (in points),
                            "font": "font name"
                        }
                        
                        Examples:

                        "move text up" would return {
        "action": "move",
        "position": "up",
        "distance": 20
    }

    "move the text down a lot" would return {
        "action": "move",
        "position": "down",
        "distance": 50
                    }
                        "make a text box with 3 bullet points about cats" would return {
                            "action": "create",
                            "type": "textbox",
                            "hasBullets": true,
                            "bulletPoints": [
                                "Cats are excellent hunters",
                                "They can sleep up to 16 hours a day",
                                "Cats have excellent night vision"
                            ]
                        }
                        
                        "create bullet points about the roman empire" would return {
                            "action": "create",
                            "type": "textbox",
                            "hasBullets": true,
                            "bulletPoints": [
                                "Founded in 27 BC by Augustus",
                                "Reached its greatest extent under Trajan",
                                "Latin was the official language",
                                "Advanced architecture including aqueducts",
                                "Powerful military system",
                                "Fell in 476 AD"
                            ]
                        }`
                    },
                    {
                        role: "user",
                        content: `Current state: ${JSON.stringify(currentState)}
                        Command: ${command}`
                    }
                ],
                temperature: 0.3
            })
        });

        if (!response.ok) {
            throw new Error(`OpenAI API error: ${response.status}`);
        }

        const data = await response.json();
        logDebug('OpenAI response received', data);

        if (!data.choices || !data.choices[0]?.message?.content) {
            throw new Error('Invalid response from OpenAI');
        }

        const instructions = JSON.parse(data.choices[0].message.content);
        logDebug('Parsed instructions', instructions);
        
        // Convert single instruction to array format for consistent processing
        return Array.isArray(instructions) ? instructions : [instructions];

    } catch (error) {
        logDebug('Error in processWithOpenAI:', error);
        return { error: 'Failed to process command with AI' };
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
        
        await shape.context.sync();
        logDebug('Formatting applied successfully');
    } catch (error) {
        logDebug('Error applying formatting:', error);
        throw new Error('Failed to apply formatting');
    }
}

async function applyFormattingToAll(slide, formatting) {
    const shapes = slide.shapes;
    shapes.load("items");
    await slide.context.sync();

    const shapeItems = shapes.items;
    let matchFound = false;
    const targetShapes = formatting.targetShapes?.toLowerCase() || "";

    for (let i = 0; i < shapeItems.length; i++) {
        const shape = shapeItems[i];
        shape.load("shapeType, textFrame, textFrame/hasText, textFrame/textRange/font, placeholder");
    }
    await slide.context.sync();

    for (let i = 0; i < shapeItems.length; i++) {
        const shape = shapeItems[i];
        let isMatch = false;

        // Determine if this shape matches the target criteria, 
        if (targetShapes.includes("all shapes") || targetShapes === "") {
            isMatch = true;
        } else if (targetShapes.includes("text boxes")) {
            if (shape.shapeType === "TextBox") {
                isMatch = true;
            }
        } else if (targetShapes.includes("titles")) {
            if (shape.placeholder && shape.placeholder.type === "Title") {
                isMatch = true;
            }
        }

        if (isMatch && shape.textFrame && shape.textFrame.hasText) {
            await applyFormatting(shape, formatting);
            matchFound = true;
        }
    }

    if (!matchFound) {
        showStatus("No matching shapes found to apply formatting.", "error");
    }
}


function createTextBox(instruction) {
    return new Promise((resolve, reject) => {
        // Format the text with bullet points if needed
        let textContent;
        if (instruction.hasBullets && instruction.bulletPoints) {
            textContent = instruction.bulletPoints.map(point => `• ${point}`).join('\n');
        } else {
            textContent = instruction.text;
        }

        Office.context.document.setSelectedDataAsync(
            textContent,
            {
                coercionType: Office.CoercionType.Text
            },
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    PowerPoint.run(async (context) => {
                        try {
                            const selection = context.presentation.getSelectedShapes();
                            selection.load("items");
                            await context.sync();

                            if (selection.items && selection.items.length > 0) {
                                const shape = selection.items[0];
                                shape.load("textFrame");
                                await context.sync();

                                const textRange = shape.textFrame.textRange;
                                textRange.load("paragraphFormat");
                                await context.sync();

                                // Center the text if not bullets
                                if (!instruction.hasBullets) {
                                    textRange.paragraphFormat.alignment = "center";
                                } else {
                                    // Left align for bullet points
                                    textRange.paragraphFormat.alignment = "left";
                                }
                                
                                // Apply formatting if it exists
                                if (instruction.formatting) {
                                    await applyFormatting(shape, instruction.formatting);
                                }

                                await context.sync();
                                logDebug('Text box created and formatted successfully');
                                resolve();
                            } else {
                                reject(new Error('No shape selected or no text box found'));
                            }
                        } catch (error) {
                            reject(error);
                        }
                    });
                } else {
                    reject(new Error(result.error.message));
                }
            }
        );
    });
}

function moveSelectedText(position, distance = 20) {
    return new Promise((resolve, reject) => {
        try {
            PowerPoint.run(async (context) => {
                // Get the selected shapes and load required properties
                const selection = context.presentation.getSelectedShapes();
                selection.load("items");
                await context.sync();

                if (selection.items && selection.items.length > 0) {
                    const shape = selection.items[0];
                    
                    // Load shape properties
                    shape.load("left,top,width");
                    await context.sync();

                    // Standard PowerPoint slide width and height
                    const slideWidth = 960;
                    const slideHeight = 540;
                    
                    let positionDescription;
                    
                    switch(position) {
                        case "left":
                            // Position at 25% (180 points)
                            shape.left = slideWidth * 0.25 - (shape.width / 2);
                            positionDescription = "left quarter";
                            break;
                        case "right":
                            // Position at 75% (540 points)
                            shape.left = slideWidth * 0.75 - (shape.width / 2);
                            positionDescription = "right quarter";
                            break;
                        case "center":
                            // Position at 50% (360 points)
                            shape.left = slideWidth * 0.5 - (shape.width / 2);
                            positionDescription = "center";
                            break;
                        case "up":
                            // Move up by distance points
                            shape.top = Math.max(0, shape.top - distance);
                            positionDescription = `up by ${distance} points`;
                            break;
                        case "down":
                            // Move down by distance points
                            shape.top = Math.min(slideHeight - 50, shape.top + distance);
                            positionDescription = `down by ${distance} points`;
                            break;
                    }
                    
                    await context.sync();
                    resolve(`Moved ${positionDescription}`);
                } else {
                    reject(new Error("No shape selected. Please select a text box first."));
                }
            }).catch(function(error) {
                reject(new Error(`Failed to move shape: ${error.message}`));
            });
        } catch (error) {
            reject(error);
        }
    });
}
async function executeCommand() {
    logDebug('executeCommand started');
    const commandInput = document.getElementById("commandInput");
    if (!commandInput) {
        logDebug('Command input not found');
        return;
    }

    const commandText = commandInput.value.trim();
    
    if (!commandText) {
        showStatus("Please enter a command", "error");
        return;
    }

    // Add user's message to chat immediately
    window.ChatApp.addMessage('user', commandText);
    showStatus("Processing with AI...", "processing");
    
    try {
        if (isImageGenerationCommand(commandText)) {
            try {
                // Extract the actual image description from the command
                const description = commandText.replace(/generate image of|create image of|make image of|draw/gi, '').trim();
                showStatus("Generating image...", "processing");
               
                // Get the base64 image
                const base64Image = await generateImageFromDescription(description);

                // Add AI's response to chat
                window.ChatApp.addMessage('assistant', "I've generated your image! Feel free to drop it down anywhere.");

                // Add the image to the last message
                const chatContainer = document.getElementById('chat-container');
                const lastMessage = chatContainer.lastElementChild;
                const messageContent = lastMessage.querySelector('.message-content');
                
                // Create image container and append to the last message
                const imageContainer = document.createElement('div');
                imageContainer.className = 'image-container';
                const img = document.createElement('img');
                img.src = `data:image/png;base64,${base64Image}`;
                img.alt = "Generated Image";
                img.className = 'generated-image';
                imageContainer.appendChild(img);
                messageContent.appendChild(imageContainer);

                // Show success status
                showStatus("Image generated successfully!", "success");
            } catch (error) {
                logDebug('Error generating image:', error);
                showStatus("Error generating image: " + error.message, "error");
                window.ChatApp.addMessage('assistant', "Sorry, I couldn't generate that image. Please try again.");
            }
        } else {

            const instructions = await processWithOpenAI(commandText, {});
            logDebug('Received instructions:', instructions);

            if (instructions.error) {
                const errorMessage = `Error: ${instructions.error}`;
                showStatus(errorMessage, "error");
                window.ChatApp.addMessage('assistant', errorMessage);
                return;
            }

            let responseMessages = [];

            // Process each instruction in sequence
            // Process each instruction in sequence
for (const instruction of instructions) {
    if (instruction.action === "create" && instruction.type === "textbox") {
        if (instruction.hasBullets && instruction.bulletPoints) {
            // Handle bullet points
            const bulletPointText = instruction.bulletPoints.map(point => `• ${point}`).join('\n');
            await createTextBox({ ...instruction, text: bulletPointText });
            responseMessages.push(`Created a bullet point list`);
        } else {
            await createTextBox(instruction);
            responseMessages.push(`Created a text box with the text "${instruction.text}"`);
        }
    } else if (instruction.action === "move") {
        const positionMessage = await moveSelectedText(
            instruction.position, 
            instruction.distance || 20
        );
        responseMessages.push(positionMessage);
    } else if (instruction.action === "format") {
        await PowerPoint.run(async (context) => {
            // First check if a shape is selected
            const selection = context.presentation.getSelectedShapes();
            selection.load("items");
            await context.sync();

            // If a shape is selected and command doesn't include "all", only format that shape
            if (selection.items && selection.items.length > 0 && !commandText.toLowerCase().includes('all')) {
                const shape = selection.items[0];
                await applyFormatting(shape, instruction);

                const changes = [];
                if (instruction.color) changes.push(`color to ${instruction.color}`);
                if (instruction.bold !== undefined) changes.push(`bold ${instruction.bold ? 'on' : 'off'}`);
                if (instruction.italic !== undefined) changes.push(`italic ${instruction.italic ? 'on' : 'off'}`);
                if (instruction.underline !== undefined) changes.push(`underline ${instruction.underline ? 'on' : 'off'}`);
                if (instruction.fontSize) changes.push(`font size to ${instruction.fontSize}pt`);
                if (instruction.font) changes.push(`font to ${instruction.font}`);

                responseMessages.push(`Changed ${changes.join(', ')}`);
            } 
            // If "all" is in the command or no shape is selected, format all shapes
            else {
                // Get the current slide
                const selectedSlides = context.presentation.getSelectedSlides();
                selectedSlides.load("items");
                await context.sync();

                let targetSlide;
                if (selectedSlides.items.length > 0) {
                    targetSlide = selectedSlides.items[0];
                } else {
                    const slides = context.presentation.slides;
                    slides.load("items");
                    await context.sync();
                    targetSlide = slides.items[0];
                }

                // Load shapes from the target slide
                const shapes = targetSlide.shapes;
                shapes.load("items");
                await context.sync();

                let appliedToAnyShape = false;
                const changes = [];

                // Apply formatting to all text-containing shapes
                for (const shape of shapes.items) {
                    try {
                        // Load specific properties for each shape
                        shape.load("textFrame");
                        await context.sync();

                        // Check if shape has a text frame
                        if (shape.textFrame) {
                            shape.textFrame.load("hasText");
                            await context.sync();

                            if (shape.textFrame.hasText) {
                                await applyFormatting(shape, instruction);
                                appliedToAnyShape = true;

                                // Build changes message only once
                                if (changes.length === 0) {
                                    if (instruction.color) changes.push(`color to ${instruction.color}`);
                                    if (instruction.bold !== undefined) changes.push(`bold ${instruction.bold ? 'on' : 'off'}`);
                                    if (instruction.italic !== undefined) changes.push(`italic ${instruction.italic ? 'on' : 'off'}`);
                                    if (instruction.underline !== undefined) changes.push(`underline ${instruction.underline ? 'on' : 'off'}`);
                                    if (instruction.fontSize) changes.push(`font size to ${instruction.fontSize}pt`);
                                    if (instruction.font) changes.push(`font to ${instruction.font}`);
                                }
                            }
                        }
                    } catch (error) {
                        // Skip shapes that can't be formatted
                        console.log('Skipping shape due to error:', error);
                        continue;
                    }
                }

                if (appliedToAnyShape) {
                    responseMessages.push(`Changed ${changes.join(', ')} for all text shapes in the slide`);
                } else {
                    throw new Error("No shapes with text found in the current slide");
                }
            }
        });
    }
}

            // Combine all response messages
            const finalResponse = responseMessages.join('. ');
            window.ChatApp.addMessage('assistant', finalResponse);
            showStatus("Actions completed successfully", "success");
        }
    } catch (error) {
        logDebug('Error in executeCommand:', error);
        const errorMessage = `Error: ${error.message}`;
        window.ChatApp.addMessage('system', errorMessage);
        showStatus(errorMessage, "error");
    }
    
    commandInput.value = "";
}


function showStatus(message, type) {
    const statusElement = document.getElementById("status");
    if (!statusElement) {
        logDebug('Status element not found');
        return;
    }

    statusElement.textContent = message;
    statusElement.className = `status-message ${type}`;
    
    if (type === "success" || type === "processing") {
        setTimeout(() => {
            statusElement.className = "status-message";
            statusElement.textContent = "";
        }, 3000);
    }
}