Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
      initialize();
  }
});

function initialize() {
  // Event listeners for 'Enter' key and 'Enter' button
  document.getElementById("send-button").addEventListener("click", sendMessage);
  document.getElementById("user-input").addEventListener("keypress", function(event) {
      if (event.key === "Enter") {
          sendMessage();
      }
  });
}

function sendMessage() {
  const userInputElement = document.getElementById("user-input");
  const userInput = userInputElement.value.trim();

  if (userInput) {
      // Display the user message in the chat container
      const chatContainer = document.getElementById("chat");
      const userMessageDiv = document.createElement("div");
      userMessageDiv.className = "user-message poppinsFont";
      userMessageDiv.textContent = userInput;
      userMessageDiv.classList.add("poppoinsFont");
      chatContainer.appendChild(userMessageDiv);


      // Clear the input field
      userInputElement.value = "";

      // Log the message to the console for debugging
      Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, (result) => {
          console.log("Sent message: ", userInput);
          if (result.status === Office.AsyncResultStatus.Failed) {
              console.error("Error logging message: " + result.error.message);
          }
      });
  }
}