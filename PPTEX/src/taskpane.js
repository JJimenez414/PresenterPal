const OPENAI_API_KEY = "Your-OpenAI-API-Key"; 

function logDebug(message, data = null) {
    console.log('Debug: ' + message, data || '');
}

// Declare these variables at the global scope
const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
let recognition;

Office.onReady(() => {
    logDebug('Office.onReady triggered');
    document.getElementById('executeCommand').onclick = executeCommand;
    document.getElementById('commandInput').addEventListener('keypress', function(event) {
        if (event.key === 'Enter') {
            executeCommand();
        }
    });

    // Speech Recognition Setup
    if (SpeechRecognition) {
        recognition = new SpeechRecognition();
        recognition.lang = 'en-US'; // Set language as needed

        document.getElementById('micButton').addEventListener('click', () => {
            if (recognition) {
                startSpeechRecognition();
            }
        });
    } else {
        console.warn('Speech Recognition API is not supported in this browser.');
        document.getElementById('micButton').disabled = true;
        document.getElementById('micButton').title = 'Speech Recognition not supported in this browser.';
    }
});

function startSpeechRecognition() {
    // Change button appearance to indicate active listening
    const micButton = document.getElementById('micButton');
    micButton.classList.add('listening');
    showStatus('Listening...', 'processing');
    recognition.start();

    recognition.onresult = (event) => {
        const transcript = event.results[0][0].transcript;
        logDebug('Speech recognition result:', transcript);

        // Set the recognized text to the command input
        document.getElementById('commandInput').value = transcript;

        // Automatically execute the command
        executeCommand();

        // Reset status and button appearance
        showStatus('', '');
        micButton.classList.remove('listening');
    };

    recognition.onerror = (event) => {
        console.error('Speech recognition error:', event.error);
        showStatus('Speech recognition error: ' + event.error, 'error');
        micButton.classList.remove('listening');
    };

    recognition.onend = () => {
        logDebug('Speech recognition ended');
        micButton.classList.remove('listening');
    };
}

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
                size: '256x256',  // Use smaller size for compatibility and performance
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

async function insertImageIntoSlide(base64Image) {
    try {
        await PowerPoint.run(async (context) => {
            // Get the active slide
            const slides = context.presentation.slides;
            slides.load('items');
            await context.sync();

            let slide;
            if (slides.items.length > 0) {
                slide = slides.items[0]; // For simplicity, use the first slide
            } else {
                throw new Error('No slides available in the presentation.');
            }

            // Insert the image from base64
            const imageShape = slide.shapes.addImageFromBase64(base64Image);
            imageShape.left = 100; // Adjust position as needed
            imageShape.top = 100;
            imageShape.height = 200; // Adjust size as needed
            imageShape.width = 200;

            await context.sync();
            logDebug('Image inserted into slide.');
        });

        showStatus('Image inserted successfully!', 'success');
    } catch (error) {
        //logDebug('Error inserting image into slide:', error);
        showStatus('Failed to insert image: ' + error.message, 'error');
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
                model: 'gpt-3.5-turbo',
                messages: [
                    {
                        role: 'system',
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
                        role: 'user',
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
        // Load necessary properties
        shape.load('textFrame/textRange/font');
        await shape.context.sync();

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
            showStatus(formatting.error, 'error');
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
    const commandText = document.getElementById('commandInput').value;

    if (!commandText.trim()) {
        showStatus('Please enter a command', 'error');
        return;
    }

    showStatus('Processing...', 'processing');

    try {
        if (isImageGenerationCommand(commandText)) {
            // Extract the actual image description from the command
            const description = commandText.replace(/generate image of|create image of|make image of|draw/gi, '').trim();
            showStatus('Generating image...', 'processing');

            // Get the base64 image
            const base64Image = await generateImageFromDescription(description);

            // Display the image in the task pane
            const testImage = document.createElement('img');
            testImage.src = `data:image/png;base64,${base64Image}`;
            testImage.alt = 'Generated Image';
            testImage.style.maxWidth = '100%';
            testImage.style.height = 'auto';
            document.getElementById('content-main').appendChild(testImage);

            // Insert the image into the slide
            await insertImageIntoSlide(base64Image);

        } else {
            // Text formatting logic
            await PowerPoint.run(async (context) => {
                let slide;
                try {
                    const selectedSlides = context.presentation.getSelectedSlides();
                    selectedSlides.load('items');
                    await context.sync();
                    slide = selectedSlides.items.length > 0 ? selectedSlides.items[0] : context.presentation.slides.getItemAt(0);
                } catch {
                    slide = context.presentation.slides.getItemAt(0);
                }

                let shape;
                try {
                    const selection = slide.shapes.getSelection();
                    selection.load('items');
                    await context.sync();
                    shape = selection.items.length > 0 ? selection.items[0] : slide.shapes.getItemAt(0);
                } catch {
                    shape = slide.shapes.getItemAt(0);
                }

                // Ensure the shape has a text frame
                shape.load('textFrame');
                await context.sync();

                if (!shape.textFrame || !shape.textFrame.hasText) {
                    showStatus('Selected shape does not contain text.', 'error');
                    return;
                }

                const currentState = await getCurrentState(shape);
                const formatting = await processWithOpenAI(commandText, currentState);
                await applyFormatting(shape, formatting);
                await context.sync();
                showStatus('Changes applied successfully', 'success');
            });
        }
    } catch (error) {
        logDebug('Error in executeCommand:', error);
        showStatus('Error: ' + error.message, 'error');
    }

    document.getElementById('commandInput').value = '';
}

function showStatus(message, type) {
    const statusElement = document.getElementById('status');
    statusElement.textContent = message;
    statusElement.className = 'status-message ' + type;

    if (type === 'success') {
        setTimeout(() => {
            statusElement.className = 'status-message';
            statusElement.textContent = '';
        }, 3000);
    }
}