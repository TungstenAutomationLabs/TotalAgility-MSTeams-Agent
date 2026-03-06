/**
 * @file utils.js
 * @module utils
 * @description Shared utility functions for the TotalAgility Teams Bot.
 *
 * Provides:
 * - Random loading/progress messages shown while the TotalAgility Agent
 *   is processing a request.
 * - Conversation history rendering (Markdown format for the Agent API,
 *   HTML format for browser-based debugging).
 */

// ── Message Pools ─────────────────────────────────────────────────────────────

/**
 * Messages shown immediately after receiving a user prompt, before the
 * TotalAgility Agent call begins.
 * @type {string[]}
 */
const loading_messages = [
  "I'm working on that for you. Please hold on a moment!",
  "Just a second, I'm getting that information for you.",
  "Hang tight! I'm processing your request.",
  "I'll have that for you shortly. Thanks for your patience!",
  "One moment please, I'm on it!",
  "Working on it! I'll be right with you.",
  "Please hold on, I'm fetching the details for you.",
  "I'm getting that information. Just a moment!",
  "I'll have an answer for you shortly. Thanks for waiting!",
  "One moment, I'm working on your request.",
  "I'm on it! Please wait a moment.",
  "Please wait while I get the details...",
  "No problem, let me look that up for you.",
];

/**
 * Messages shown at periodic intervals while the TotalAgility Agent call
 * is still in progress (long-running requests).
 * @type {string[]}
 */
const progress_messages = [
  "Performing research...",
  "Running agents...",
  "Gathering results...",
  "Analysing results.",
  "Planning next action.",
  "Thinking...",
  "Processing data...",
  "Sorry for the wait, I'm still working on it!",
  "Apologies for the delay, this is taking a bit longer than expected.",
];

// ── Public Functions ──────────────────────────────────────────────────────────

/**
 * Render a conversation history array as an HTML DOM tree.
 *
 * > **Note:** This function uses `document.createElement` and is intended for
 * > browser-based debugging only.  It will throw in a Node.js environment
 * > unless a DOM polyfill is available.
 *
 * @param {string[]} conversationArray - Alternating user/bot messages.
 * @deprecated Use {@link renderConversationHistoryMarkdown} for server-side rendering.
 */
function renderConversationHistory(conversationArray) {
  const conversationContainer = document.createElement("div");
  conversationContainer.style.border = "1px solid #ccc";
  conversationContainer.style.padding = "10px";
  conversationContainer.style.maxWidth = "600px";
  conversationContainer.style.margin = "20px auto";
  conversationContainer.style.fontFamily = "Arial, sans-serif";

  for (let i = 0; i < conversationArray.length; i++) {
    const message = document.createElement("div");
    message.style.marginBottom = "10px";

    if (i % 2 === 0) {
      message.style.color = "blue";
      message.innerHTML = `<strong>User:</strong> ${conversationArray[i]}`;
    } else {
      message.style.color = "green";
      message.innerHTML = `<strong>Bot:</strong> ${conversationArray[i]}`;
    }

    conversationContainer.appendChild(message);
  }

  document.body.appendChild(conversationContainer);
}

/**
 * Render a conversation history array as a Markdown-formatted string.
 *
 * This is the format sent to the TotalAgility Agent so it has full context
 * of the conversation when generating a response.
 *
 * @param {Array<{speaker: string, message: string}>} conversationArray
 *   Each element must have `speaker` (e.g. "User") and `message` properties.
 * @returns {string} Markdown representation of the conversation.
 * @throws {TypeError} If `conversationArray` is not an array.
 *
 * @example
 * const md = renderConversationHistoryMarkdown([
 *   { speaker: "User", message: "What's in my workqueue?" },
 *   { speaker: "TotalAgility Bot", message: "You have 3 items." },
 * ]);
 * // Returns:
 * // **User:** What's in my workqueue?
 * //
 * // **TotalAgility Bot:** You have 3 items.
 */
function renderConversationHistoryMarkdown(conversationArray) {
  if (!Array.isArray(conversationArray)) {
    throw new TypeError("Input must be an array");
  }

  if (conversationArray.length === 0) {
    return "No conversation history available.";
  }

  const formattedConversation = conversationArray.map((item, index) => {
    if (typeof item !== "object" || !item.speaker || !item.message) {
      return `[Invalid entry at position ${index}]`;
    }
    return `**${item.speaker}:** ${item.message.trim()} \n\n`;
  });

  return formattedConversation.join("\n\n");
}

/**
 * Pick a random loading message to show the user while their request is
 * being submitted.
 *
 * @returns {string} A randomly selected loading message.
 */
function getRandomLoadingMessage() {
  const randomIndex = Math.floor(Math.random() * loading_messages.length);
  return loading_messages[randomIndex];
}

/**
 * Pick a random progress message to show the user during long-running
 * TotalAgility Agent calls.
 *
 * @returns {string} A randomly selected progress message.
 */
function getRandomProgressMessage() {
  const randomIndex = Math.floor(Math.random() * progress_messages.length);
  return progress_messages[randomIndex];
}

// ── Config Debugging ──────────────────────────────────────────────────────────

/**
 * Keys whose values must never appear in logs or user-visible output.
 * @type {Set<string>}
 */
const SENSITIVE_CONFIG_KEYS = new Set([
  "MicrosoftAppPassword",
  "totalAgilityApiKey",
  "notificationsBearerToken",
  "azureStorageConnectionString",
]);

/**
 * Build a human-readable summary of the current configuration.
 *
 * Sensitive values (API keys, passwords, tokens, connection strings) are
 * masked — only whether they are set or not is shown.
 *
 * @param {Object} configObj - The config object (from `require("./config")`).
 * @returns {{ lines: string[], markdown: string }}
 *   `lines`    — plain-text lines suitable for `console.log()`.
 *   `markdown` — Markdown-formatted string suitable for sending to the user.
 */
function getConfigSummary(configObj) {
  const lines = [];
  const mdLines = [];

  for (const [key, value] of Object.entries(configObj)) {
    let display;
    if (SENSITIVE_CONFIG_KEYS.has(key)) {
      display = value ? "******** (set)" : "(not set)";
    } else {
      display =
        value === undefined || value === null || value === ""
          ? "(not set)"
          : value;
    }
    lines.push(`${key} = ${display}`);
    mdLines.push(`**${key}:** \`${display}\``);
  }

  return {
    lines,
    markdown: mdLines.join("  \n"),
  };
}

// ── Exports ───────────────────────────────────────────────────────────────────
module.exports = {
  renderConversationHistory,
  renderConversationHistoryMarkdown,
  getRandomLoadingMessage,
  getRandomProgressMessage,
  getConfigSummary,
};
