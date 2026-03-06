/**
 * @file teamsBot.js
 * @module TeamsBot
 * @description Microsoft Teams bot handler for the TotalAgility Agent.
 *
 * Extends the Bot Framework's `TeamsActivityHandler` to:
 * - Receive user messages and forward them (with conversation history) to the
 *   TotalAgility Agent REST API.
 * - Handle file attachments uploaded by the user (converted to base64).
 * - Show typing indicators and periodic "still working" messages for
 *   long-running agent calls.
 * - Capture and persist `ConversationReference` objects so that proactive
 *   notifications can be delivered later via `/api/notifications`.
 *
 * @see {@link module:taAgent}   TotalAgility API integration
 * @see {@link module:conversationStore}  Conversation reference persistence
 * @see {@link module:utils}     Helper utilities (loading messages, history rendering)
 */

// ── Dependencies ──────────────────────────────────────────────────────────────
const { TeamsActivityHandler, TurnContext, TeamsInfo } = require("botbuilder");
const TotalAgilityAgent = require("./taAgent.js");
const config = require("./config");
const Utils = require("./utils.js");
const conversationStore = require("./conversationStore.js");
const axios = require("axios");

// ── Module-Level State ────────────────────────────────────────────────────────

/**
 * Rolling conversation history sent to the TotalAgility Agent on each turn.
 * Each entry is `{ speaker: string, message: string }`.
 * @type {Array<{speaker: string, message: string}>}
 */
let messageArray = [];

/**
 * Maximum number of messages retained in `messageArray`.
 * Configurable via the `CONVERSATION_HISTORY_MAX_ENTRIES` environment variable.
 * @type {number}
 * @default 10
 */
const messageArrayMaxSize = (() => {
  const val = parseInt(config.conversationHistoryMaxEntries, 10);
  return isNaN(val) || val <= 0 ? 10 : val;
})();

/** Current TotalAgility SSO session key (refreshed on every message). */
let ssoKey = "";

/**
 * In-memory cache mapping AAD object IDs to email addresses.
 * Avoids repeated `TeamsInfo.getMember()` HTTP round-trips.
 * @type {Map<string, string>}
 */
const emailCache = new Map();

/**
 * Millisecond intervals at which periodic "still working" messages are sent
 * while awaiting a TotalAgility Agent response.
 * @type {number[]}
 */
const PROGRESS_INTERVALS = [
  15000, 22000, 30000, 40000, 50000, 60000, 70000, 80000, 90000, 100000,
  110000, 120000, 130000, 140000, 150000, 160000, 170000, 180000, 190000,
  200000,
];

/**
 * Lookup table mapping file extensions to MIME types.
 * Used when forwarding uploaded files to TotalAgility.
 * @type {Object<string, string>}
 */
const MIME_TYPES = {
  jpg: "image/jpeg",
  jpeg: "image/jpeg",
  png: "image/png",
  gif: "image/gif",
  bmp: "image/bmp",
  webp: "image/webp",
  svg: "image/svg+xml",
  ico: "image/x-icon",
  tiff: "image/tiff",
  tif: "image/tiff",
  pdf: "application/pdf",
  txt: "text/plain",
  html: "text/html",
  htm: "text/html",
  css: "text/css",
  js: "application/javascript",
  json: "application/json",
  xml: "application/xml",
  csv: "text/csv",
  zip: "application/zip",
  rar: "application/vnd.rar",
  tar: "application/x-tar",
  gz: "application/gzip",
  mp3: "audio/mpeg",
  wav: "audio/wav",
  ogg: "audio/ogg",
  mp4: "video/mp4",
  avi: "video/x-msvideo",
  mov: "video/quicktime",
  wmv: "video/x-ms-wmv",
  flv: "video/x-flv",
  mkv: "video/x-matroska",
  doc: "application/msword",
  docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  xls: "application/vnd.ms-excel",
  xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  ppt: "application/vnd.ms-powerpoint",
  pptx: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
};

// ── TeamsBot Class ────────────────────────────────────────────────────────────

/**
 * Bot that handles incoming Teams activities and forwards user messages to the
 * TotalAgility Agent.
 *
 * @extends TeamsActivityHandler
 */
class TeamsBot extends TeamsActivityHandler {
  /**
   * Create a new TeamsBot instance.
   *
   * @param {import("botbuilder").UserState} userState - Bot Framework UserState for per-user persistence.
   * @param {import("botbuilder").StatePropertyAccessor} ssoKeyAccessor - State accessor for the TotalAgility SSO key.
   */
  constructor(userState, ssoKeyAccessor) {
    super();
    this.userState = userState;
    this.ssoKeyAccessor = ssoKeyAccessor;

    // ── Message Handler ─────────────────────────────────────────────────
    this.onMessage(async (context, next) => {
      try {
        // Store the conversation reference for proactive notifications.
        // Fire-and-forget so it doesn't block the message processing path.
        this._saveConversationReference(context).catch((err) =>
          console.error("[TeamsBot] Background save failed:", err.message)
        );

        // Guard: context.activity.text can be null for attachment-only messages.
        const messageText = (context.activity.text || "")
          .toLowerCase()
          .replace(/\n|\r/g, "")
          .trim();

        if (messageText === "debug") {
          // ── Debug command ────────────────────────────────────────────
          // Prints all loaded config values (sensitive values masked) to
          // both the console log and the user's Teams chat.
          const { lines, markdown } = Utils.getConfigSummary(config);
          console.log("[TeamsBot] Debug command invoked — config dump:");
          lines.forEach((line) => console.log(`[Config]   ${line}`));
          await context.sendActivity(
            `**🔧 Configuration (sensitive values masked):**\n\n${markdown}`
          );
        } else if (
          messageText.match(
            /^(clear conversation history|clear history|clear|reset|clear conversation)$/
          )
        ) {
          await context.sendActivity(
            "Current conversation history: " +
              Utils.renderConversationHistoryMarkdown(messageArray)
          );
          messageArray = [];
          await context.sendActivity("Conversation history reset.");
        } else {
          // Send an initial loading message and typing indicator.
          await context.sendActivity(Utils.getRandomLoadingMessage());
          await context.sendActivities([{ type: "typing" }]);

          // Authenticate with TotalAgility via SSO.
          try {
            ssoKey = await TotalAgilityAgent.taSSOLogin(context);
          } catch (ssoErr) {
            console.error("[TeamsBot] SSO login failed:", ssoErr.message);
            await context.sendActivity(
              `⚠️ Unable to sign into TotalAgility. Please try again in a moment.\n\nError: ${ssoErr.message}`
            );
            return;
          }

          await context.sendActivities([{ type: "typing" }]);
          await this.handleMessageWithLoadingIndicator(context, ssoKey);
          await next();
        }
      } catch (err) {
        console.error("[TeamsBot] Unexpected error in onMessage:", err);
        try {
          await context.sendActivity(
            `⚠️ Sorry, something went wrong while processing your message.\n\nError: ${err.message}`
          );
        } catch (_) {
          console.error(
            "[TeamsBot] Failed to send error message to user:",
            _.message
          );
        }
      }
    });

    // ── Members Added Handler ───────────────────────────────────────────
    // Fires when users are added to a conversation (including bot installation).
    // See: https://docs.microsoft.com/en-us/microsoftteams/platform/resources/bot-v3/bots-notifications
    this.onMembersAdded(async (context, next) => {
      const membersAdded = context.activity.membersAdded;
      for (let cnt = 0; cnt < membersAdded.length; cnt++) {
        if (membersAdded[cnt].id) {
          // Capture the conversation reference on install so proactive
          // notifications work even before the user sends their first message.
          await this._saveConversationReference(context);
          break;
        }
      }
      await next();
    });

    // ── Conversation Update Handler ─────────────────────────────────────
    // Fires on events such as the bot being added to a group chat.
    this.onConversationUpdate(async (context, next) => {
      await this._saveConversationReference(context);
      await next();
    });
  }

  // ── Private Methods ───────────────────────────────────────────────────────

  /**
   * Store the current user's `ConversationReference` in the conversation store
   * so that proactive messages can be sent later via `/api/notifications`.
   *
   * **User key resolution** (matches the SSO logic in `taAgent.js`):
   * 1. If `TOTALAGILITY_USE_TEST_USER` is `"true"` → uses `TOTALAGILITY_TEST_USERNAME`.
   * 2. Otherwise → resolves the Teams user's email via `TeamsInfo.getMember()`
   *    (cached in `emailCache` to avoid repeated HTTP calls).
   * 3. Falls back to the user's display name or Bot Framework ID.
   *
   * @param {import("botbuilder").TurnContext} context - The current turn context.
   * @returns {Promise<void>}
   * @private
   */
  async _saveConversationReference(context) {
    try {
      const ref = TurnContext.getConversationReference(context.activity);
      let userKey = null;

      if (config.totalAgilityUseTestUser === "true") {
        userKey = config.totalAgilityTestUserName;
      } else {
        const aadId =
          context.activity.from && context.activity.from.aadObjectId;
        if (aadId) {
          if (emailCache.has(aadId)) {
            userKey = emailCache.get(aadId);
          } else {
            try {
              const member = await TeamsInfo.getMember(
                context,
                context.activity.from.id
              );
              if (member && member.email) {
                userKey = member.email;
                emailCache.set(aadId, userKey);
              }
            } catch (_) {
              // TeamsInfo may not be available (e.g. during conversationUpdate).
            }
          }
        }
        if (!userKey && context.activity.from) {
          userKey = context.activity.from.name || context.activity.from.id;
        }
      }

      if (userKey) {
        await conversationStore.save(userKey, ref);
        console.log("[TeamsBot] Saved conversation reference for:", userKey);
      }
    } catch (err) {
      console.error(
        "[TeamsBot] Error saving conversation reference:",
        err.message
      );
    }
  }

  /**
   * Process a user message: download any file attachments, call the
   * TotalAgility Agent API, and return the response to the user.
   *
   * While waiting for the agent response, periodic "still working" messages
   * are sent at the intervals defined in `PROGRESS_INTERVALS`.
   *
   * @param {import("botbuilder").TurnContext} context - The current turn context.
   * @param {string} ssoKey - The TotalAgility SSO session key.
   * @returns {Promise<void>}
   */
  async handleMessageWithLoadingIndicator(context, ssoKey) {
    await context.sendActivities([{ type: "typing" }]);
    console.log("Running with Message Activity.");

    try {
      // ── Extract & normalise user text ───────────────────────────────
      const removedMentionText = TurnContext.removeRecipientMention(
        context.activity
      );
      const userRequest = removedMentionText
        .toLowerCase()
        .replace(/\n|\r/g, "")
        .trim();
      saveMsg("User", userRequest);

      // ── File attachment handling ────────────────────────────────────
      let base64String = "";
      let mimeType = "";
      let fileName = "";
      let documentId = "";

      if (
        context.activity.attachments &&
        context.activity.attachments.length > 0
      ) {
        for (let i = 0; i < context.activity.attachments.length; i++) {
          const attachment = context.activity.attachments[i];

          if (
            attachment.contentType ===
            "application/vnd.microsoft.teams.file.download.info"
          ) {
            const downloadUrl = attachment.content.downloadUrl;
            fileName = attachment.name;
            mimeType = getMimeType(attachment.content.fileType);

            await context.sendActivity(`Uploading file ${fileName}.`);
            await context.sendActivities([{ type: "typing" }]);

            // Download and convert to base64 for the TotalAgility API.
            const response = await axios.get(downloadUrl, {
              responseType: "arraybuffer",
            });
            base64String = Buffer.from(response.data).toString("base64");

            await context.sendActivity(`File ${fileName} received.`);
            await context.sendActivities([{ type: "typing" }]);

            // ── Document preload (optional) ──────────────────────────
            // When PRELOAD_DOCUMENTS_AS_TOTALAGILITY_DOCS is enabled,
            // submit the file to a dedicated Document Creator process
            // to obtain a TotalAgility Document ID.  This avoids storing
            // the large base64 string in the process database.
            if (config.preloadDocumentsAsTotalAgilityDocs === "true" && base64String) {
              await context.sendActivity("Creating TotalAgility document...");
              await context.sendActivities([{ type: "typing" }]);

              documentId = await TotalAgilityAgent.createTotalAgilityDocument(
                base64String,
                mimeType,
                ssoKey,
                fileName
              );

              if (documentId) {
                console.log("[TeamsBot] Document preloaded, ID:", documentId);
                // Clear all file content fields — only the Document ID will be
                // sent to the Chat Agent via the DOCUMENT input variable.
                base64String = "";
                mimeType = "";
                fileName = "";
                await context.sendActivity(
                  `✅ TotalAgility Document created: \`${documentId}\``
                );
                await context.sendActivities([{ type: "typing" }]);
              } else {
                console.warn(
                  "[TeamsBot] Document preload failed — falling back to inline base64."
                );
              }
            }
          }
        }
      } else {
        await context.sendActivity(
          "No file attached. Processing your message..."
        );
      }

      // ── Progress timers ─────────────────────────────────────────────
      let timers = [];
      let isDone = false;

      PROGRESS_INTERVALS.forEach((delay, idx) => {
        timers[idx] = setTimeout(async () => {
          if (!isDone) {
            await context.sendActivity(Utils.getRandomProgressMessage());
            await context.sendActivities([{ type: "typing" }]);
          }
        }, delay);
      });

      // ── Call the TotalAgility Agent ─────────────────────────────────
      const agentResponse = await TotalAgilityAgent.callRestService(
        Utils.renderConversationHistoryMarkdown(messageArray),
        base64String,
        mimeType,
        ssoKey,
        fileName,
        documentId
      );
      await context.sendActivity(agentResponse);

      // ── Cleanup ─────────────────────────────────────────────────────
      isDone = true;
      timers.forEach((timer) => clearTimeout(timer));
      saveMsg("TotalAgility Bot", agentResponse);
    } catch (error) {
      await context.sendActivity(
        `⚠️ An error occurred: ${error.message}`
      );
    }
  }
}

// ── Module-Level Helper Functions ─────────────────────────────────────────────

/**
 * Append a message to the rolling conversation history.
 * When the array exceeds `messageArrayMaxSize`, the oldest entry is removed.
 *
 * @param {string} actor   - The speaker label (e.g. "User", "TotalAgility Bot").
 * @param {string} message - The message text.
 */
function saveMsg(actor, message) {
  messageArray.push({ speaker: actor, message: message });
  if (messageArray.length > messageArrayMaxSize) {
    messageArray.shift();
  }
}

/**
 * Map a file extension to its MIME type using the `MIME_TYPES` lookup table.
 *
 * @param {string} ext - File extension (with or without leading dot).
 * @returns {string|null} The MIME type, or `null` if unrecognised.
 *
 * @example
 * getMimeType('pdf');   // → 'application/pdf'
 * getMimeType('.DOCX'); // → 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
 * getMimeType('xyz');   // → null
 */
function getMimeType(ext) {
  ext = ext.replace(/^\./, "").toLowerCase();
  return MIME_TYPES[ext] || null;
}

module.exports.TeamsBot = TeamsBot;
