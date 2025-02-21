const config = {
  MicrosoftAppId: process.env.BOT_ID,
  MicrosoftAppType: process.env.BOT_TYPE,
  MicrosoftAppTenantId: process.env.BOT_TENANT_ID,
  MicrosoftAppPassword: process.env.BOT_PASSWORD,
  totalAgilityEndpoint: process.env.TOTALAGILITY_ENDPOINT,
  totalAgilityApiKey: process.env.TOTALAGILITY_API_KEY,
  totalAgilityAgentName: process.env.TOTALAGILITY_AGENT_NAME,
  totalAgilityAgentId: process.env.TOTALAGILITY_AGENT_ID,
};

module.exports = config;
