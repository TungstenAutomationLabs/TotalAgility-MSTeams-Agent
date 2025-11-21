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

// Gracefully shutdown HTTP server
["exit", "uncaughtException", "SIGINT", "SIGTERM", "SIGUSR1", "SIGUSR2"].forEach((event) => {
  process.on(event, () => {
    server.close();
  });
});
