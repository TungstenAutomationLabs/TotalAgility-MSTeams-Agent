/**
 * @file taAgent.js
 * @module taAgent
 * @description TotalAgility REST API integration layer.
 *
 * Provides functions to:
 * - Authenticate a user into TotalAgility via SSO (`taSSOLogin`).
 * - Submit a prompt (with optional file attachment) to a TotalAgility Agent
 *   process and retrieve the agent's response (`callRestService`).
 * - Optionally pre-create a TotalAgility Document from an uploaded file
 *   (`createTotalAgilityDocument`) so that only a lightweight document ID
 *   is passed to the chat agent — avoiding large base64 strings in the
 *   process database.
 *
 * The TotalAgility "sync" job endpoint (`/jobs/sync`) is used so that the
 * process runs in-memory on the server for lower latency.
 *
 * **Document handling modes** (controlled by `PRELOAD_DOCUMENTS_AS_TOTALAGILITY_DOCS`):
 * - `false` (default) — the raw base64 string, MIME type, and filename are
 *   sent as input variables (`DOCUMENT_CONTENT`, `DOCUMENT_TYPE`,
 *   `DOCUMENT_FILENAME`) directly to the chat agent process, with
 *   `DOCUMENT` left empty.
 * - `true` — the file is first submitted to a separate "Document Creator"
 *   process which stores it in TotalAgility's document storage and returns
 *   a Document ID.  That ID is then passed to the chat agent via the
 *   `DOCUMENT` input variable, while `DOCUMENT_CONTENT`, `DOCUMENT_TYPE`,
 *   and `DOCUMENT_FILENAME` are sent as empty strings — reducing database
 *   load significantly.
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
 * Pre-create a TotalAgility Document by submitting the file to a dedicated
 * "Document Creator" process via `/jobs/sync`.
 *
 * This is the **recommended** approach for production deployments because it
 * stores the document once in TotalAgility's document storage and returns a
 * lightweight Document ID.  The chat agent can then reference the document by
 * ID rather than receiving a large base64 string as a process variable, which
 * reduces load on the TotalAgility process database.
 *
 * The payload uses the TotalAgility `Documents` array to submit the file
 * content directly as a document attachment (with `Base64Data`), rather than
 * passing it as a process input variable.  The response returns the
 * `DocumentId` at the top level of the JSON body.
 *
 * Required config values (set via environment variables):
 * - `totalAgilityDocumentCreatorProcessId`  — Process ID (GUID)
 * - `totalAgilityDocumentCreatorProcessName` — Process name
 * - `totalAgilityDocumentTypeId`            — Document Type ID (GUID)
 * - `totalAgilityDocumentFilenameFieldId`   — RuntimeField ID for filename
 *
 * @param {string} base64String  - Base64-encoded file content.
 * @param {string} mimeType      - MIME type of the file (e.g. `"application/pdf"`).
 * @param {string} sessionKey    - TotalAgility SSO session ID.
 * @param {string} fileName      - Original filename (e.g. `"report.pdf"`).
 * @returns {Promise<string>} The TotalAgility Document ID, or empty string on failure.
 */
async function createTotalAgilityDocument(base64String, mimeType, sessionKey, fileName) {
  console.log("[taAgent] createTotalAgilityDocument() called for:", fileName);

  if (!base64String) return "";

  const url = config.totalAgilityEndpoint + "/jobs/sync";

  const payload = {
    ProcessId: config.totalAgilityDocumentCreatorProcessId,
    ProcessName: config.totalAgilityDocumentCreatorProcessName,
    JobInitialization: {
      InputVariables: [],
    },
    Documents: [
      {
        MimeType: mimeType || "",
        RuntimeFields: [
          {
            Id: config.totalAgilityDocumentFilenameFieldId || "",
            TableRow: -1,
            TableColumn: -1,
            Value: fileName || "",
          },
        ],
        FolderId: "",
        DocumentTypeId: config.totalAgilityDocumentTypeId || "",
        FolderTypeId: "",
        Base64Data: base64String,
        DocumentTypeName: "",
        DocumentGroupId: "",
        DocumentGroupName: "",
      },
    ],
    VariablesToReturn: [],
    StoreFolderAndDocuments: true,
    ReturnOnlySpecifiedDocuments: true,
  };

  const headers = {
    "Content-Type": "application/json",
    Authorization: sessionKey,
  };

  try {
    const response = await fetch(url, {
      method: "POST",
      headers: headers,
      body: JSON.stringify(payload),
    });

    if (!response.ok) {
      throw new Error(
        `Document Creator HTTP error! status: ${response.status}, URL: ${url}`
      );
    }

    const data = await response.json();
    console.log("[taAgent] Document Creator JobId:", data.JobId);

    // The top-level DocumentId in the response is the created document's ID
    if (data.DocumentId) {
      console.log("[taAgent] Created TotalAgility Document ID:", data.DocumentId);
      return data.DocumentId;
    }

    console.warn("[taAgent] Document Creator did not return a DocumentId.");
    return "";
  } catch (error) {
    console.error("[taAgent] Document Creator error:", error);
    return "";
  }
}

/**
 * Call the TotalAgility Agent via the `/jobs/sync` REST endpoint.
 *
 * Builds a job payload containing the user's prompt (with conversation
 * history), optional file attachment, and LLM tuning parameters,
 * then POSTs it to TotalAgility and extracts the `OUTPUT` variable from the
 * response.
 *
 * **Document handling:**
 * Document-related input variables (`DOCUMENT`, `DOCUMENT_CONTENT`,
 * `DOCUMENT_TYPE`, `DOCUMENT_FILENAME`) are **only included in the payload
 * when the user explicitly attaches a file** in the current turn.  This
 * prevents the same document from being re-sent on every subsequent message.
 *
 * - If `documentInfo.documentId` is provided (preload mode), it is sent as
 *   `DOCUMENT` and the content/type/filename fields are empty.
 * - If `documentInfo` contains `base64String` (inline mode), the raw content,
 *   MIME type, and filename are sent as `DOCUMENT_CONTENT`, `DOCUMENT_TYPE`,
 *   and `DOCUMENT_FILENAME`, with `DOCUMENT` left empty.
 * - If `documentInfo` is `null` / `undefined`, **no document variables are
 *   included** in the payload at all.
 *
 * @param {string} prompt_text   - The full conversation history rendered as Markdown.
 * @param {string} sessionKey    - TotalAgility SSO session ID (used as `Authorization` header).
 * @param {Object|null} [documentInfo=null] - Document attachment for this turn, or null/undefined if none.
 * @param {string} [documentInfo.documentId]  - TotalAgility Document ID (preload mode).
 * @param {string} [documentInfo.base64String] - Base64-encoded file content (inline mode).
 * @param {string} [documentInfo.mimeType]     - MIME type of the file (inline mode).
 * @param {string} [documentInfo.fileName]     - Original filename (inline mode).
 * @returns {Promise<string>} The agent's text response, or an error message string.
 */
async function callRestService(prompt_text, sessionKey, documentInfo) {
  console.log("callRestService() called with: " + prompt_text);

  // Destructure document info only when provided.
  const hasDocument = !!documentInfo;
  const documentId   = (documentInfo && documentInfo.documentId)   || "";
  const base64String = (documentInfo && documentInfo.base64String) || "";
  const mimeType     = (documentInfo && documentInfo.mimeType)     || "";
  const fileName     = (documentInfo && documentInfo.fileName)     || "";

  if (hasDocument) {
    if (documentId) {
      console.log("[taAgent] Using preloaded TotalAgility Document ID:", documentId);
    } else if (base64String) {
      console.log("[taAgent] File attached with size:", base64String.length);
    }
    if (mimeType) console.log("[taAgent] File MIME type:", mimeType);
  } else {
    console.log("[taAgent] No document attached for this turn.");
  }

  let return_response = "";
  const url = config.totalAgilityEndpoint + "/jobs/sync";
  console.log("Calling TotalAgility on " + url);

  // ── Build the job payload ─────────────────────────────────────────────
  const inputVariables = [
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
  ];

  // ── Document variables (only when a file was attached this turn) ─────
  // When no file is attached, these variables are omitted entirely so the
  // TotalAgility process receives no stale document data.
  if (hasDocument) {
    if (documentId) {
      // Preload mode — pass only the document reference.
      inputVariables.push({ Id: "DOCUMENT_TYPE", Value: "" });
      inputVariables.push({ Id: "DOCUMENT_CONTENT", Value: "" });
      inputVariables.push({ Id: "DOCUMENT_FILENAME", Value: "" });
      inputVariables.push({ Id: "DOCUMENT", Value: documentId });
    } else {
      // Inline mode — pass the raw base64 content.
      inputVariables.push({ Id: "DOCUMENT_TYPE", Value: mimeType });
      inputVariables.push({ Id: "DOCUMENT_CONTENT", Value: base64String });
      inputVariables.push({ Id: "DOCUMENT_FILENAME", Value: fileName });
      inputVariables.push({ Id: "DOCUMENT", Value: "" });
    }
  }

  const payload = {
    ProcessId: config.totalAgilityAgentId,
    ProcessName: config.totalAgilityAgentName,
    JobInitialization: {
      InputVariables: inputVariables,
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
  createTotalAgilityDocument,
  taSSOLogin,
};
