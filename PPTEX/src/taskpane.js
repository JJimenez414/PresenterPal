// Replace your API key securely (see note below)
function logDebug(message, data = null) {
    console.log(`Debug: ${message}`, data || '');
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


async function processWithOpenAI(command, currentState) {
    // (No changes in this function)
    // ... existing code ...
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
                        content: `You are a PowerPoint formatting assistant. Convert natural language commands into specific formatting instructions.
Respond only with a JSON object containing the following possible properties:
{
    "color": "color name", // Accept color names like 'red', 'blue', etc.
    "bold": boolean,
    "italic": boolean,
    "underline": boolean,
    "fontSize": number (in points),
    "font": "font name",
    "applyToAll": boolean, // Set to true if the command refers to multiple shapes
    "targetShapes": "description of target shapes", // e.g., "text boxes", "titles", "all shapes"
    "error": "error message if command is invalid"
}
Examples:
1. "Make all text boxes red and bold" would return {"color": "red", "bold": true, "applyToAll": true, "targetShapes": "text boxes"}
2. "Make all titles italic and blue" would return {"color": "blue", "italic": true, "applyToAll": true, "targetShapes": "titles"}
3. "Change the color to green for all text, or all text boxes" would return {"color": "green", "applyToAll": true, "targetShapes": "all shapes"}
Always return color names instead of hex codes.`
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

        const formattingInstructions = JSON.parse(data.choices[0].message.content);
        logDebug('Parsed formatting instructions', formattingInstructions);
        return formattingInstructions;

    } catch (error) {
        logDebug('Error in processWithOpenAI:', error);
        return { error: 'Failed to process command with AI' };
    }
}

async function getCurrentState(shape) {
    // (No changes in this function)
    // ... existing code ...
    try {
        const state = {
            color: shape.textFrame.textRange.font.color || 'black',
            bold: shape.textFrame.textRange.font.bold || false,
            italic: shape.textFrame.textRange.font.italic || false,
            underline: shape.textFrame.textRange.font.underline || false,
            fontSize: shape.textFrame.textRange.font.size || 12,
            font: shape.textFrame.textRange.font.name || 'Calibri'
        };
        logDebug('Current state retrieved', state);
        return state;
    } catch (error) {
        logDebug('Error getting current state:', error);
        return {
            color: 'black',
            bold: false,
            italic: false,
            underline: false,
            fontSize: 12,
            font: 'Calibri'
        };
    }
}

async function applyFormatting(shape, formatting) {
    // (No changes in this function)
    // ... existing code ...
    try {
        logDebug('Applying formatting', formatting);
        
        if (formatting.error) {
            showStatus(formatting.error, "error");
            return;
        }

        const font = shape.textFrame.textRange.font;

        if (formatting.color) {
            font.color = formatting.color; // Accept color names directly
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

async function executeCommand() {
    logDebug('executeCommand started');
    const commandText = document.getElementById("commandInput").value;
    
    if (!commandText.trim()) {
        showStatus("Please enter a command", "error");
        return;
    }

    showStatus("Processing with AI...", "processing");
    
    try {
        await PowerPoint.run(async (context) => {
            // Load all slides in the presentation
            const slides = context.presentation.slides;
            slides.load("items");
            await context.sync();

            // For AI context, get current state from the first shape of the first slide
            let currentState = {};
            let foundState = false;

            for (let slideIndex = 0; slideIndex < slides.items.length; slideIndex++) {
                const slide = slides.items[slideIndex];
                const shapes = slide.shapes;
                shapes.load("items");
                await context.sync();

                for (let i = 0; i < shapes.items.length; i++) {
                    const shape = shapes.items[i];
                    shape.load("textFrame, textFrame/hasText, textFrame/textRange/font");
                }
                await context.sync();

                for (let i = 0; i < shapes.items.length; i++) {
                    const shape = shapes.items[i];
                    if (shape.textFrame && shape.textFrame.hasText) {
                        currentState = await getCurrentState(shape);
                        foundState = true;
                        break;
                    }
                }

                if (foundState) {
                    break;
                }
            }

            if (!foundState) {
                showStatus("No shapes with text found to get current state.", "error");
                return;
            }

            const formatting = await processWithOpenAI(commandText, currentState);

            if (formatting.error) {
                showStatus(formatting.error, "error");
                return;
            }

            // Determine the target slide
            let targetSlideIndex = null; // 0-based index
            if (formatting.targetSlide !== undefined && formatting.targetSlide !== null) {
                targetSlideIndex = formatting.targetSlide - 1; // Convert to 0-based index
                if (targetSlideIndex < 0 || targetSlideIndex >= slides.items.length) {
                    showStatus(`Slide ${formatting.targetSlide} does not exist.`, "error");
                    return;
                }
            } else {
                // If no targetSlide specified, use the current slide
                const selectedSlides = context.presentation.getSelectedSlides();
                selectedSlides.load("items");
                await context.sync();

                if (selectedSlides.items.length > 0) {
                    targetSlideIndex = slides.items.findIndex(s => s.id === selectedSlides.items[0].id);
                } else {
                    // Fallback to the first slide
                    targetSlideIndex = 0;
                }
            }

            // Get the target slide
            const targetSlide = slides.items[targetSlideIndex];

            // Apply formatting to the target slide
            const shapes = targetSlide.shapes;
            shapes.load("items");
            await context.sync();

            for (let i = 0; i < shapes.items.length; i++) {
                const shape = shapes.items[i];
                shape.load("textFrame, textFrame/hasText, textFrame/textRange/font, shapeType, placeholder");
            }
            await context.sync();

            if (formatting.applyToAll) {
                await applyFormattingToAll(targetSlide, formatting);
            } else {
                // Apply to the first matching shape on the slide
                let shapeApplied = false;
                for (let i = 0; i < shapes.items.length; i++) {
                    const shape = shapes.items[i];
                    if (shape.textFrame && shape.textFrame.hasText) {
                        await applyFormatting(shape, formatting);
                        shapeApplied = true;
                        break;
                    }
                }

                if (!shapeApplied) {
                    logDebug(`No matching shape found on slide ${targetSlideIndex + 1}`);
                }
            }

            await context.sync();

            showStatus("Changes applied successfully", "success");
        });
    } catch (error) {
        logDebug('Error in executeCommand:', error);
        showStatus(`Error: ${error.message}`, "error");
    }

    document.getElementById("commandInput").value = "";
}

function showStatus(message, type) {
    const statusElement = document.getElementById("status");
    statusElement.textContent = message;
    statusElement.className = `status-message ${type}`;
    
    if (type === "success") {
        setTimeout(() => {
            statusElement.className = "status-message";
        }, 3000);
    }
}