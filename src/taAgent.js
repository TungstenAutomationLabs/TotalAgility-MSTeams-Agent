/**
 * @file taAgent.js
 * @module taAgent
 * @description TotalAgility REST API integration layer.
 *
 * Provides functions to:
 * - Authenticate a user into TotalAgility via SSO (`taSSOLogin`).
 * - Submit a prompt (with optional file attachment) to a TotalAgility Agent
 *   process and retrieve the agent's response (`callRestService`).
 *
 * The TotalAgility "sync" job endpoint (`/jobs/sync`) is used so that the
 * process runs in-memory on the server for lower latency.  Documents are
 * passed as base64-encoded input variables rather than via the `Documents`
 * array, to avoid a race condition where the document hasn't been persisted
 * by the time sub-jobs try to access it.
 *
 * @see {@link module:config} for environment-driven configuration values.
 */

const config = require("./config");
const { TeamsInfo } = require("botbuilder");

// ── Public API ────────────────────────────────────────────────────────────────

/**
 * Simple connectivity test (development only).
 *
 * @param {string} text - Arbitrary text to echo back.
 * @returns {string} Greeting string.
 */
function tester(text) {
  console.log("Calling TotalAgility on " + config.totalAgilityEndpoint);
  console.log("Calling TotalAgility on " + config.totalAgilityApiKey);
  return "Hello " + text;
}

/**
 * Call the TotalAgility Agent via the `/jobs/sync` REST endpoint.
 *
 * Builds a job payload containing the user's prompt (with conversation
 * history), optional file attachment (base64), and LLM tuning parameters,
 * then POSTs it to TotalAgility and extracts the `OUTPUT` variable from the
 * response.
 *
 * @param {string} prompt_text   - The full conversation history rendered as Markdown.
 * @param {string} base64String  - Base64-encoded file content (empty string if no file).
 * @param {string} mimeType      - MIME type of the attached file (empty string if no file).
 * @param {string} sessionKey    - TotalAgility SSO session ID (used as `Authorization` header).
 * @param {string} fileName      - Original filename of the attachment (empty string if no file).
 * @returns {Promise<string>} The agent's text response, or an error message string.
 */
async function callRestService(prompt_text, base64String, mimeType, sessionKey, fileName) {
  console.log("callRestService() called with: " + prompt_text);

  // Normalise optional parameters to empty strings.
  if (!base64String) base64String = "";
  if (!mimeType) mimeType = "";
  if (!fileName) fileName = "";

  if (base64String) console.log("File attached with size: " + base64String.length);
  if (mimeType) console.log("File MIME type: " + mimeType);

  let return_response = "";
  const url = config.totalAgilityEndpoint + "/jobs/sync";
  console.log("Calling TotalAgility on " + url);

  // ── Build the job payload ─────────────────────────────────────────────
  const payload = {
    ProcessId: config.totalAgilityAgentId,
    ProcessName: config.totalAgilityAgentName,
    JobInitialization: {
      InputVariables: [
        {
          Id: "INPUT_PROMPT",
          Value: prompt_text,
        },
        {
          Id: "TEMPERATURE",
          Value: (() => {
            const t = parseFloat(config.totalAgilityTemperature);
            return isNaN(t) ? 1 : t;
          })(),
        },
        {
          Id: "USE_SEED",
          Value: (() => {
            const u = config.totalAgilityUseSeed;
            if (typeof u === "string") return u.toLowerCase() === "true";
            return u === undefined ? true : !!u;
          })(),
        },
        {
          Id: "SEED",
          Value: (() => {
            const s = parseInt(config.totalAgilitySeed, 10);
            return isNaN(s) ? 27535 : s;
          })(),
        },
        { Id: "DOCUMENT_CONTENT", Value: base64String },
        { Id: "DOCUMENT_TYPE", Value: mimeType },
        { Id: "DOCUMENT_FILENAME", Value: fileName },
      ],
    },
    Documents: [],
    VariablesToReturn: [{ VarId: "OUTPUT" }],
    StoreFolderAndDocuments: true,
    ReturnOnlySpecifiedDocuments: true,
  };

  const headers = {
    "Content-Type": "application/json",
    Authorization: sessionKey,
  };

  // ── Execute the request ───────────────────────────────────────────────
  try {
    const response = await fetch(url, {
      method: "POST",
      headers: headers,
      body: JSON.stringify(payload),
    });

    if (!response.ok) {
      throw new Error(
        `HTTP error! status: ${response.status} \n\n URL: ${url} \n\n`
      );
    }

    const data = await response.json();
    console.log("JobId:", data.JobId);

    // Extract the OUTPUT returned variable.
    if (data.ReturnedVariables && data.ReturnedVariables.length > 0) {
      data.ReturnedVariables.forEach((variable) => {
        console.log("Returned Variable Value:", variable.Value);
        return_response = variable.Value;
      });
    } else {
      console.log("No Returned Variables found.");
      return_response = "No Returned Variables found.";
    }
  } catch (error) {
    console.error("Error: ", error);
    return_response =
      "Error: " + error + "\nPayload: " + JSON.stringify(payload);
  }

  return return_response;
}

/**
 * Authenticate a user into TotalAgility via the SSO endpoint.
 *
 * In **test mode** (`TOTALAGILITY_USE_TEST_USER=true`), authenticates as the
 * configured test user.  In **production mode**, resolves the current Teams
 * user's email address and uses it as the TotalAgility User ID.
 *
 * The TotalAgility SSO API returns an existing session if one is already
 * active, so it is safe (and recommended) to call this on every message turn.
 *
 * @param {import("botbuilder").TurnContext} context - The current bot turn context.
 * @returns {Promise<string>} The TotalAgility session ID.
 * @throws {Error} If the SSO HTTP request fails.
 */
async function taSSOLogin(context) {
  try {
    const ssoUrl =
      config.totalAgilityEndpoint + "/users/sessions/single-sign-on";

    const ssoPayload = { UserId: "" };

    if (config.totalAgilityUseTestUser === "true") {
      ssoPayload.UserId = config.totalAgilityTestUserName;
    } else {
      const userInfo = await getCurrentUserIdAndEmail(context);
      ssoPayload.UserId = userInfo.email;
    }

    const ssoHeaders = {
      "Content-Type": "application/json",
      Authorization: config.totalAgilityApiKey,
    };

    const ssoResponse = await fetch(ssoUrl, {
      method: "POST",
      headers: ssoHeaders,
      body: JSON.stringify(ssoPayload),
    });

    if (!ssoResponse.ok) {
      throw new Error(
        `HTTP error! status: ${ssoResponse.status} \n\n URL: ${ssoUrl}` +
          ` \n\n Use Test user: ${config.totalAgilityUseTestUser}` +
          ` \n\n Test UserID: ${config.totalAgilityTestUserName}` +
          ` \n\n Payload: ${JSON.stringify(ssoPayload)}`
      );
    }

    const ssoData = await ssoResponse.json();
    return ssoData.SessionId;
  } catch (error) {
    console.error("SSO Login Error: ", error);
    throw error; // Bubble up — caught in teamsBot.js onMessage handler.
  }
}

// ── Private Helpers ───────────────────────────────────────────────────────────

/**
 * Retrieve the current Teams user's ID and email address from the Bot
 * Framework service.
 *
 * @param {import("botbuilder").TurnContext} context - The current bot turn context.
 * @returns {Promise<{id: string, email: string}|null>} User info, or `null` on failure.
 * @private
 */
async function getCurrentUserIdAndEmail(context) {
  try {
    const member = await TeamsInfo.getMember(
      context,
      context.activity.from.id
    );
    return { id: member.id, email: member.email };
  } catch (error) {
    console.error("Failed to get user info:", error);
    return null;
  }
}

// ── Exports ───────────────────────────────────────────────────────────────────
module.exports = {
  tester,
  callRestService,
  taSSOLogin,
};
