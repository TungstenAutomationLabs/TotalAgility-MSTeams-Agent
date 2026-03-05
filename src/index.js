// index.js is used to setup and configure your bot

// Import required packages
const express = require("express");

// Import required bot services.
// See https://aka.ms/bot-services to learn more about the different parts of a bot.
const {
  CloudAdapter,
  ConfigurationServiceClientCredentialFactory,
  ConfigurationBotFrameworkAuthentication,
  UserState,
  MemoryStorage
} = require("botbuilder");

const { TeamsBot } = require("./teamsBot");
const conversationStore = require("./conversationStore");

const config = require("./config");

// How state across messages, so only one SSO key is needed for TA.
// const { UserState, MemoryStorage } = require("botbuilder");

// Create adapter.
// See https://aka.ms/about-bot-adapter to learn more about adapters.
const credentialsFactory = new ConfigurationServiceClientCredentialFactory(config);

const botFrameworkAuthentication = new ConfigurationBotFrameworkAuthentication(
  {},
  credentialsFactory
);


const adapter = new CloudAdapter(botFrameworkAuthentication);

// Set up state management for SSO key storage:
const memoryStorage = new MemoryStorage();
const userState = new UserState(memoryStorage);
const ssoKeyAccessor = userState.createProperty('ssoKey');

adapter.onTurnError = async (context, error) => {
  // This check writes out errors to console log .vs. app insights.
  // NOTE: In production environment, you should consider logging this to Azure
  //       application insights. See https://aka.ms/bottelemetry for telemetry
  //       configuration instructions.
  console.error(`\n [onTurnError] unhandled error: ${error}`);

  // Only send error message for user messages, not for other message types so the bot doesn't spam a channel or chat.
  if (context.activity.type === "message") {
    // Send a message to the user
    await context.sendActivity(`The bot encountered an unhandled error:\n ${error.message}`);
    await context.sendActivity("To continue to run this bot, please fix the bot source code.");
  }
};

// Create the bot that will handle incoming messages.
//const bot = new TeamsBot();
// Create the bot that will handle incoming messages & maintain user state for SSO key: 
const bot = new TeamsBot(userState, ssoKeyAccessor);

// Create express application.
const expressApp = express();
expressApp.use(express.json());

// Initialise the conversation reference store (Azure Table Storage or in-memory fallback).
conversationStore.init().then(() => {
  console.log("[index] Conversation store ready.");
}).catch((err) => {
  console.error("[index] Conversation store init error:", err.message);
});

const server = expressApp.listen(process.env.port || process.env.PORT || 3978, () => {
  console.log(`\nBot Started, ${expressApp.name} listening to`, server.address());
  //console.log("TotalAgility Endpoint: " + config.totalAgilityEndpoint);
  //console.log("TotalAgility API Key: " + config.totalAgilityApiKey);
  //console.log("TotalAgility Agent Name: " + config.totalAgilityAgentName);
  //console.log("TotalAgility Agent ID: " + config.totalAgilityAgentId);
});

// Listen for incoming requests.
expressApp.post("/api/messages", async (req, res) => {
  await adapter.process(req, res, async (context) => {
    await bot.run(context);
    await userState.saveChanges(context); // Save any state changes.
  });
});

// ─────────────────────────────────────────────────────────────────────────────
//  Middleware: Bearer-token authentication for notification endpoints
// ─────────────────────────────────────────────────────────────────────────────
function requireNotificationAuth(req, res, next) {
  const expectedToken = config.notificationsBearerToken;
  if (!expectedToken) {
    return res.status(503).json({
      error: "Notification endpoint is disabled — NOTIFICATIONS_BEARER_TOKEN is not configured.",
    });
  }

  const authHeader = req.headers.authorization || "";
  const token = authHeader.startsWith("Bearer ") ? authHeader.slice(7) : authHeader;

  if (!token || token !== expectedToken) {
    return res.status(401).json({ error: "Unauthorized — invalid or missing bearer token." });
  }

  next();
}

// ─────────────────────────────────────────────────────────────────────────────
//  POST /api/notifications  — send a proactive message to a named user
// ─────────────────────────────────────────────────────────────────────────────
//
//  Request body (JSON):
//  {
//    "userKey":  "jane.doe@contoso.com",   // email or name used during registration
//    "message":  "Your document has been processed."
//  }
//
//  The user must have previously interacted with the bot (or had the bot
//  installed) so that a ConversationReference exists in the store.
//
expressApp.post("/api/notifications", requireNotificationAuth, async (req, res) => {
  try {
    const { userKey, message } = req.body || {};

    if (!userKey || !message) {
      return res.status(400).json({
        error: "Request body must include 'userKey' (string) and 'message' (string).",
      });
    }

    const ref = await conversationStore.get(userKey);
    if (!ref) {
      return res.status(404).json({
        error: `No conversation reference found for user '${userKey}'. ` +
               "The user must message the bot at least once before they can receive notifications.",
      });
    }

    // Send the proactive message using the stored ConversationReference.
    await adapter.continueConversationAsync(
      config.MicrosoftAppId,
      ref,
      async (turnContext) => {
        await turnContext.sendActivity(message);
      }
    );

    return res.status(200).json({ status: "ok", userKey, message });
  } catch (err) {
    console.error("[/api/notifications] Error:", err);
    return res.status(500).json({ error: err.message });
  }
});

// ─────────────────────────────────────────────────────────────────────────────
//  GET /api/conversations  — list all registered user keys (diagnostic)
// ─────────────────────────────────────────────────────────────────────────────
expressApp.get("/api/conversations", requireNotificationAuth, async (_req, res) => {
  try {
    const all = await conversationStore.listAll();
    return res.status(200).json({ count: all.length, conversations: all });
  } catch (err) {
    console.error("[/api/conversations] Error:", err);
    return res.status(500).json({ error: err.message });
  }
});

// Gracefully shutdown HTTP server
["exit", "uncaughtException", "SIGINT", "SIGTERM", "SIGUSR1", "SIGUSR2"].forEach((event) => {
  process.on(event, () => {
    server.close();
  });
});
