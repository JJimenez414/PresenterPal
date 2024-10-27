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
                            "color": "hex color code",
                            "bold": boolean,
                            "italic": boolean,
                            "underline": boolean,
                            "fontSize": number (in points),
                            "error": "error message if command is invalid"
                        }
                        Example 1: "make it red and bold" would return {"color": "#FF0000", "bold": true}
                        Example 2: "remove bold and make it blue" would return {"color": "#0000FF", "bold": false}
                        Always return valid hex codes for colors.`
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
    try {
        const state = {
            color: shape.textFrame.textRange.font.color || '#000000',
            bold: shape.textFrame.textRange.font.bold || false,
            italic: shape.textFrame.textRange.font.italic || false,
            underline: shape.textFrame.textRange.font.underline || false,
            fontSize: shape.textFrame.textRange.font.size || 12
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
            fontSize: 12
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
        
        logDebug('Formatting applied successfully');
    } catch (error) {
        logDebug('Error applying formatting:', error);
        throw new Error('Failed to apply formatting');
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
            let slide;
            try {
                slide = context.presentation.getSelectedSlides();
                await context.sync();
                slide = slide.items.length > 0 ? slide.items[0] : context.presentation.slides.getItemAt(0);
            } catch {
                slide = context.presentation.slides.getItemAt(0);
            }
            
            let shape;
            try {
                const selection = slide.shapes.getSelection();
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