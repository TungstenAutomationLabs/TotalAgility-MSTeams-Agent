
const loading_messages = [
    "I'm working on that for you. Please hold on a moment!",
    "Just a second, I'm getting that information for you.",
    "Hang tight! I'm processing your request.",
    "I'll have that for you shortly. Thanks for your patience!",
    "One moment please, I'm on it!",
    "I'm on the job! Please wait a moment.",
    "Working on it! I'll be right with you.",
    "Please hold on, I'm fetching the details for you.",
    "I'm getting that information. Just a moment!",
    "I'll have an answer for you shortly. Thanks for waiting!",
    "Give me a moment, I'm working on your request.",
    "I'm on it! Please wait a moment.",
    "Please wait while I get the details...",
    "No problem, let me look that up for you."
  ];

const progress_messages = [
    "Still working on it, thanks for your patience!",
    "I'm making progress, please hold on a bit longer.",    
    "Thanks for waiting, I'm still on it!",
    "Performing research...",
    "Running agents...",
    "Gathering results...",
    "Analysing results.",
    "Planning next action.",
    "Thinking...",
    "Appreciate your patience, I'm still working on your request.",
    "I'm getting closer to the answer, thanks for bearing with me!",
    "Still processing your request, thank you for your understanding.",
    "I'm still on the job, thanks for sticking with me!",
    "Making headway, please hold on a little longer.",
    "Sorry for the wait, I'm still working on it!",
    "Apologies for the delay, this is taking a bit longer than expected."
];


function renderConversationHistory(conversationArray) {
    // Create a container for the conversation
    const conversationContainer = document.createElement('div');
    conversationContainer.style.border = '1px solid #ccc';
    conversationContainer.style.padding = '10px';
    conversationContainer.style.maxWidth = '600px';
    conversationContainer.style.margin = '20px auto';
    conversationContainer.style.fontFamily = 'Arial, sans-serif';

    // Loop through the conversation array
    for (let i = 0; i < conversationArray.length; i++) {
        const message = document.createElement('div');
        message.style.marginBottom = '10px';

        // Alternate styles for user and bot messages
        if (i % 2 === 0) {
            message.style.color = 'blue'; // User messages in blue
            message.innerHTML = `<strong>User:</strong> ${conversationArray[i]}`;
        } else {
            message.style.color = 'green'; // Bot responses in green
            message.innerHTML = `<strong>Bot:</strong> ${conversationArray[i]}`;
        }

        // Append the message to the conversation container
        conversationContainer.appendChild(message);
    }

    // Append the conversation container to the body or a specific element
    document.body.appendChild(conversationContainer);
}


/**
 * Renders an array of text elements as a conversation history in Markdown format.
 * @param {Array} conversationArray - An array of objects representing the conversation.
 * @returns {string} A formatted string representing the conversation history in Markdown.
 */
function renderConversationHistoryMarkdown(conversationArray) {
    // Input validation
    if (!Array.isArray(conversationArray)) {
        throw new TypeError("Input must be an array");
    }

    // Handle empty array
    if (conversationArray.length === 0) {
        return "No conversation history available.";
    }

    // Use map() to transform each conversation element into a formatted Markdown string
    const formattedConversation = conversationArray.map((item, index) => {
        // Ensure each item is an object with 'speaker' and 'message' properties
        if (typeof item !== 'object' || !item.speaker || !item.message) {
            return `[Invalid entry at position ${index}]`;
        }

        // Use Markdown formatting
        return `**${item.speaker}:** ${item.message.trim()} \n\n`;
    });

    // Join the formatted strings with newlines
    return formattedConversation.join('\n\n');
}

function getRandomLoadingMessage() {
    const randomIndex = Math.floor(Math.random() * loading_messages.length);
    return loading_messages[randomIndex];
}

function getRandomProgressMessage() {
    const randomIndex = Math.floor(Math.random() * progress_messages.length);
    return progress_messages[randomIndex];
}

// Export the functions
module.exports = {
    renderConversationHistory,
    renderConversationHistoryMarkdown,
    getRandomLoadingMessage,
    getRandomProgressMessage
};