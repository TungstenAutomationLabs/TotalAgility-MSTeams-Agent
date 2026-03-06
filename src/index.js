/**
 * @file index.js
 * @module index
 * @description Application entry point for the TotalAgility Teams Bot.
 *
 * Bootstraps the Express HTTP server, configures the Bot Framework adapter,
 * initialises the conversation reference store, and exposes the following
 * endpoints:
 *
 * | Method | Path                | Purpose                                         |
 * |--------|---------------------|-------------------------------------------------|
 * | POST   | `/api/messages`     | Bot Framework messaging endpoint (Teams ↔ Bot)  |
 * | POST   | `/api/notifications`| Proactive notification endpoint (3rd-party → user) |
 * | GET    | `/api/conversations`| List registered conversation references           |
 *
 * @see {@link https://aka.ms/bot-services} Bot Framework overview
 * @see {@link https://learn.microsoft.com/en-us/microsoftteams/platform/bots/how-to/conversations/send-proactive-messages} Proactive messaging guide
 */

// ── Node.js crypto polyfill ───────────────────────────────────────────────────
// Some Azure App Service Node.js runtimes don't expose `globalThis.crypto`
// (required by @typespec/ts-http-runtime for UUID generation).  Polyfill it
// from the built-in `node:crypto` module before any other imports.
if (!globalThis.crypto) {
  globalThis.crypto = require("crypto");
}

// ── Dependencies ──────────────────────────────────────────────────────────────
const express = require("express");
const helmet = require("helmet");
const rateLimit = require("express-rate-limit");
const {
  CloudAdapter,
  ConfigurationServiceClientCredentialFactory,
  ConfigurationBotFrameworkAuthentication,
  UserState,
  MemoryStorage,
} = require("botbuilder");

const { TeamsBot } = require("./teamsBot");
const conversationStore = require("./conversationStore");
const config = require("./config");
const Utils = require("./utils");

// ── Startup Config Validation ─────────────────────────────────────────────────
/**
 * Validate that critical configuration values are present at startup.
 * Logs warnings for missing optional values and errors for missing required ones.
 */
(function validateConfig() {
  const required = [
    ["TOTALAGILITY_ENDPOINT", config.totalAgilityEndpoint],
    ["TOTALAGILITY_API_KEY", config.totalAgilityApiKey],
    ["TOTALAGILITY_AGENT_NAME", config.totalAgilityAgentName],
    ["TOTALAGILITY_AGENT_ID", config.totalAgilityAgentId],
  ];

  const optional = [
    ["BOT_ID", config.MicrosoftAppId],
    ["NOTIFICATIONS_BEARER_TOKEN", config.notificationsBearerToken],
    ["AZURE_STORAGE_CONNECTION_STRING", config.azureStorageConnectionString],
  ];

  let hasErrors = false;
  for (const [name, value] of required) {
    if (!value) {
      console.error(`[Config] ❌ Missing required env var: ${name}`);
      hasErrors = true;
    }
  }

  for (const [name, value] of optional) {
    if (!value) {
      console.warn(`[Config] ⚠️  Missing optional env var: ${name}`);
    }
  }

  if (hasErrors) {
    console.error(
      "[Config] One or more required environment variables are missing. " +
        "The bot may not function correctly."
    );
  }
})();

// ── Log Loaded Configuration ──────────────────────────────────────────────────
// Logs all configuration key-value pairs at startup so operators can verify
// that the correct env files are being picked up.  Sensitive values (API keys,
// passwords, tokens, connection strings) are masked.
(function logConfig() {
  const { lines } = Utils.getConfigSummary(config);
  console.log("[Config] ── Loaded configuration ──────────────────────────");
  lines.forEach((line) => console.log(`[Config]   ${line}`));
  console.log("[Config] ─────────────────────────────────────────────────");
})();

// ── Bot Framework Adapter ─────────────────────────────────────────────────────
// The adapter translates incoming HTTP requests into Bot Framework Activities
// and routes them to the bot logic.
// See: https://aka.ms/about-bot-adapter

const credentialsFactory = new ConfigurationServiceClientCredentialFactory(config);
const botFrameworkAuthentication = new ConfigurationBotFrameworkAuthentication(
  {},
  credentialsFactory
);
const adapter = new CloudAdapter(botFrameworkAuthentication);

// ── User State (SSO key persistence) ─────────────────────────────────────────
// MemoryStorage is used to persist user-scoped state (e.g. the TotalAgility
// SSO session key) across conversation turns within the same process.
// NOTE: MemoryStorage is not durable across restarts.  For production at
// scale, consider Azure Blob Storage or Cosmos DB.

const memoryStorage = new MemoryStorage();
const userState = new UserState(memoryStorage);
const ssoKeyAccessor = userState.createProperty("ssoKey");

// ── Global Turn Error Handler ─────────────────────────────────────────────────
/**
 * Catch-all for unhandled errors during a bot turn.
 * Logs the error and sends a user-friendly message (for message activities only).
 *
 * @param {import("botbuilder").TurnContext} context - The current turn context.
 * @param {Error} error - The unhandled error.
 */
adapter.onTurnError = async (context, error) => {
  console.error(`\n [onTurnError] unhandled error: ${error}`);

  // Only send error message for user messages — avoids spamming channels.
  if (context.activity.type === "message") {
    await context.sendActivity(
      `The bot encountered an unhandled error:\n ${error.message}`
    );
    await context.sendActivity(
      "To continue to run this bot, please fix the bot source code."
    );
  }
};

// ── Bot Instance ──────────────────────────────────────────────────────────────
const bot = new TeamsBot(userState, ssoKeyAccessor);

// ── Express Application ───────────────────────────────────────────────────────
const expressApp = express();

// ── Security Middleware ───────────────────────────────────────────────────────
// Helmet sets various HTTP headers to help protect the app from well-known
// web vulnerabilities (XSS, clickjacking, MIME sniffing, etc.).
expressApp.use(helmet());

// Limit JSON body size to 1 MB to prevent oversized payloads.
expressApp.use(express.json({ limit: "1mb" }));

/**
 * Rate limiter for notification and conversation-listing endpoints.
 * Allows 60 requests per minute per IP.  Adjust if your 3rd-party
 * integration sends bursts of notifications.
 */
const notificationLimiter = rateLimit({
  windowMs: 60 * 1000, // 1 minute
  max: 60,
  standardHeaders: true,
  legacyHeaders: false,
  message: { error: "Too many requests — please try again later." },
});

// Initialise the conversation reference store (Azure Table Storage or
// in-memory fallback).  This is fire-and-forget at startup; failures are
// logged but don't prevent the server from starting.
conversationStore
  .init()
  .then(() => console.log("[index] Conversation store ready."))
  .catch((err) =>
    console.error("[index] Conversation store init error:", err.message)
  );

// Start listening.
const server = expressApp.listen(
  process.env.port || process.env.PORT || 3978,
  () => {
    console.log(
      `\nBot Started, ${expressApp.name} listening to`,
      server.address()
    );
  }
);

// ── POST /api/messages ────────────────────────────────────────────────────────
/**
 * Main Bot Framework messaging endpoint.
 * Receives activities from the Bot Framework Channel Service and forwards them
 * to the bot's turn handler.
 */
expressApp.post("/api/messages", async (req, res) => {
  try {
    await adapter.process(req, res, async (context) => {
      await bot.run(context);
      await userState.saveChanges(context);
    });
  } catch (err) {
    console.error("[/api/messages] Unhandled error:", err);
    if (!res.headersSent) {
      res.status(500).json({ error: "Internal server error processing message." });
    }
  }
});

// ── Bearer-Token Auth Middleware ──────────────────────────────────────────────
/**
 * Express middleware that validates the `Authorization: Bearer <token>` header
 * against the configured `NOTIFICATIONS_BEARER_TOKEN`.
 *
 * Applied to the `/api/notifications` and `/api/conversations` endpoints to
 * prevent unauthorised access from external callers.
 *
 * Uses timing-safe comparison to prevent timing attacks on the bearer token.
 *
 * @param {import("express").Request}  req  - Express request object.
 * @param {import("express").Response} res  - Express response object.
 * @param {import("express").NextFunction} next - Express next middleware.
 */
function requireNotificationAuth(req, res, next) {
  const expectedToken = config.notificationsBearerToken;
  if (!expectedToken) {
    return res.status(503).json({
      error:
        "Notification endpoint is disabled — NOTIFICATIONS_BEARER_TOKEN is not configured.",
    });
  }

  const authHeader = req.headers.authorization || "";
  const token = authHeader.startsWith("Bearer ")
    ? authHeader.slice(7)
    : authHeader;

  if (!token || !timingSafeEqual(token, expectedToken)) {
    return res
      .status(401)
      .json({ error: "Unauthorized — invalid or missing bearer token." });
  }

  next();
}

/**
 * Constant-time string comparison to prevent timing attacks.
 * Falls back to simple equality if `crypto.timingSafeEqual` is unavailable.
 *
 * @param {string} a - First string.
 * @param {string} b - Second string.
 * @returns {boolean} Whether the strings are equal.
 * @private
 */
function timingSafeEqual(a, b) {
  try {
    const crypto = require("crypto");
    const bufA = Buffer.from(a);
    const bufB = Buffer.from(b);
    // If lengths differ, still perform comparison on equal-length buffers
    // to avoid leaking length information.
    if (bufA.length !== bufB.length) {
      crypto.timingSafeEqual(bufA, bufA); // constant-time no-op
      return false;
    }
    return crypto.timingSafeEqual(bufA, bufB);
  } catch {
    return a === b;
  }
}

/** Maximum allowed length for a notification message (in characters). */
const MAX_NOTIFICATION_MESSAGE_LENGTH = 4000;

// ── POST /api/notifications ──────────────────────────────────────────────────
/**
 * Proactive notification endpoint.
 *
 * Allows 3rd-party systems (e.g. TotalAgility workflows, Power Automate,
 * external APIs) to send a message directly into a user's Teams chat with the
 * bot.  Uses the Microsoft-recommended `adapter.continueConversationAsync()`
 * pattern with a stored `ConversationReference`.
 *
 * @route POST /api/notifications
 *
 * @example Request body:
 * {
 *   "userKey": "jane.doe@contoso.com",
 *   "message": "Your document has been processed."
 * }
 *
 * @example Success response (200):
 * { "status": "ok", "userKey": "jane.doe@contoso.com", "message": "..." }
 *
 * @example Error response (404):
 * { "error": "No conversation reference found for user 'jane.doe@contoso.com'..." }
 */
expressApp.post(
  "/api/notifications",
  notificationLimiter,
  requireNotificationAuth,
  async (req, res) => {
    try {
      const { userKey, message } = req.body || {};

      if (!userKey || !message) {
        return res.status(400).json({
          error:
            "Request body must include 'userKey' (string) and 'message' (string).",
        });
      }

      if (typeof userKey !== "string" || typeof message !== "string") {
        return res.status(400).json({
          error: "'userKey' and 'message' must be strings.",
        });
      }

      if (message.length > MAX_NOTIFICATION_MESSAGE_LENGTH) {
        return res.status(400).json({
          error: `'message' must not exceed ${MAX_NOTIFICATION_MESSAGE_LENGTH} characters (received ${message.length}).`,
        });
      }

      const ref = await conversationStore.get(userKey);
      if (!ref) {
        return res.status(404).json({
          error:
            `No conversation reference found for user '${userKey}'. ` +
            "The user must message the bot at least once before they can receive notifications.",
        });
      }

      // Resume the conversation and deliver the proactive message.
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
  }
);

// ── GET /api/conversations ───────────────────────────────────────────────────
/**
 * Diagnostic endpoint that lists all users with stored conversation references.
 * Useful for discovering valid `userKey` values for the notifications endpoint.
 *
 * @route GET /api/conversations
 *
 * @example Success response (200):
 * {
 *   "count": 2,
 *   "conversations": [
 *     { "userKey": "jane.doe@contoso.com", "conversationId": "...", "userName": "Jane Doe" }
 *   ]
 * }
 */
expressApp.get(
  "/api/conversations",
  notificationLimiter,
  requireNotificationAuth,
  async (_req, res) => {
    try {
      const all = await conversationStore.listAll();
      return res.status(200).json({ count: all.length, conversations: all });
    } catch (err) {
      console.error("[/api/conversations] Error:", err);
      return res.status(500).json({ error: err.message });
    }
  }
);

// ── Global Express Error Handler ─────────────────────────────────────────────
/**
 * Catch-all error handler for Express routes.
 * Prevents unhandled errors from crashing the process.
 */
expressApp.use((err, _req, res, _next) => {
  console.error("[Express] Unhandled error:", err);
  if (!res.headersSent) {
    res.status(500).json({ error: "An unexpected error occurred." });
  }
});

// ── Graceful Shutdown ─────────────────────────────────────────────────────────
["exit", "uncaughtException", "SIGINT", "SIGTERM", "SIGUSR1", "SIGUSR2"].forEach(
  (event) => {
    process.on(event, (reason) => {
      if (reason) {
        console.error(`[Process] ${event}:`, reason);
      }
      server.close();
    });
  }
);

// Prevent silent failures from unhandled promise rejections.
process.on("unhandledRejection", (reason, _promise) => {
  console.error("[Process] Unhandled promise rejection:", reason);
});
