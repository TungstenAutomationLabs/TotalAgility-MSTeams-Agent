const config = {
  MicrosoftAppId: process.env.BOT_ID,
  MicrosoftAppType: process.env.BOT_TYPE,
  MicrosoftAppTenantId: process.env.BOT_TENANT_ID,
  MicrosoftAppPassword: process.env.BOT_PASSWORD,
  totalAgilityEndpoint: process.env.TOTALAGILITY_ENDPOINT,
  totalAgilityApiKey: process.env.TOTALAGILITY_API_KEY,
  totalAgilityAgentName: process.env.TOTALAGILITY_AGENT_NAME,
  totalAgilityAgentId: process.env.TOTALAGILITY_AGENT_ID,
  totalAgilityTestUserName: process.env.TOTALAGILITY_TEST_USERNAME,
  totalAgilityUseTestUser: process.env.TOTALAGILITY_USE_TEST_USER,
  // optional LLM parameters (strings). Defaults applied in code.
  totalAgilityTemperature: process.env.TOTALAGILITY_TEMPERATURE,
  totalAgilityUseSeed: process.env.TOTALAGILITY_USE_SEED,
  totalAgilitySeed: process.env.TOTALAGILITY_SEED,
  // Conversation history
  conversationHistoryMaxEntries: process.env.CONVERSATION_HISTORY_MAX_ENTRIES,
  // Proactive notifications
  notificationsBearerToken: process.env.NOTIFICATIONS_BEARER_TOKEN,
  // Azure Table Storage for persisting conversation references
  // (used by the proactive notification endpoint).
  // When absent the store falls back to in-memory.
  azureStorageConnectionString: process.env.AZURE_STORAGE_CONNECTION_STRING,
};

module.exports = config;
