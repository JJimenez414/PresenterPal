// Create a global object for chat functions
(function(window) {
    // Store chat messages in memory
    let chatHistory = [{
        type: 'assistant',
        content: 'Hello! I can help you format your PowerPoint slides. Try commands like "make the text blue" or "increase font size".'
    }];

    // Save chat history to localStorage
    function saveChatHistory() {
        try {
            localStorage.setItem('powerpoint-chat-history', JSON.stringify(chatHistory));
            console.log("Chat history saved", chatHistory);
        } catch (error) {
            console.error("Error saving chat history:", error);
        }
    }

    // Load chat history from localStorage
    function loadChatHistory() {
        try {
            const saved = localStorage.getItem('powerpoint-chat-history');
            if (saved) {
                chatHistory = JSON.parse(saved);
                console.log("Chat history loaded", chatHistory);
            }
        } catch (error) {
            console.error("Error loading chat history:", error);
        }
    }

    // Render all messages
    function renderChatHistory() {
        const chatContainer = document.querySelector('.chat-container');
        if (!chatContainer) {
            console.error("Chat container not found");
            return;
        }

        chatContainer.innerHTML = '';
        
        chatHistory.forEach(message => {
            const messageDiv = document.createElement('div');
            messageDiv.className = `message ${message.type}`;

            const avatar = document.createElement('div');
            avatar.className = `avatar ${message.type}-avatar`;
            avatar.textContent = message.type === 'user' ? 'U' : 'AI';

            const content = document.createElement('div');
            content.className = 'message-content';
            content.textContent = message.content;

            if (message.type === 'user') {
                messageDiv.appendChild(content);
                messageDiv.appendChild(avatar);
            } else {
                messageDiv.appendChild(avatar);
                messageDiv.appendChild(content);
            }

            chatContainer.appendChild(messageDiv);
        });

        // Scroll to bottom
        chatContainer.scrollTo({
            top: chatContainer.scrollHeight,
            behavior: 'smooth'
        });
    }

    // Add new message
    function addMessage(type, content) {
        chatHistory.push({ type, content });
        saveChatHistory();
        renderChatHistory();
    }

    // Initialize
    function initialize() {
        console.log("Initializing chat history...");
        loadChatHistory();
        renderChatHistory();
    }

    // Attach all functions to the global ChatApp namespace
    window.ChatApp = {
        initialize: initialize,
        addMessage: addMessage,
        renderChatHistory: renderChatHistory,
        loadChatHistory: loadChatHistory,
        saveChatHistory: saveChatHistory
    };

    // Log that ChatApp is ready
    console.log("ChatApp initialized:", window.ChatApp);
})(window);