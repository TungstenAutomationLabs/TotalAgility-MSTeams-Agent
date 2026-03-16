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
 * **Session management** (added in v1.8):
 * SSO session keys are cached per user and reused across requests. If the
 * TotalAgility API returns a 403 "Invalid Session ID" error, the module
 * automatically clears the stale session, requests a fresh one via SSO, and
 * retries the failed call exactly once. A second consecutive 403 is treated
 * as a permanent authentication failure and returns a user-friendly error
 * message — preventing infinite retry loops.
 *
 * **Performance optimisations** (added in v1.9):
 * - **HTTP connection pooling** — a shared `undici.Pool` keeps TCP/TLS
 *   connections alive to the TotalAgility origin, eliminating the TLS
 *   handshake overhead on every request.
 * - **Request timeouts** — all HTTP calls use `AbortController` with
 *   configurable per-call-type timeouts (`TOTALAGILITY_SSO_TIMEOUT_MS`,
 *   `TOTALAGILITY_AGENT_TIMEOUT_MS`, `TOTALAGILITY_DOCUMENT_TIMEOUT_MS`).
 * - **SSO login deduplication** — concurrent session refresh requests for
 *   the same user are coalesced into a single in-flight SSO call, preventing
 *   a thundering-herd of SSO logins.
 * - **Timing instrumentation** — every HTTP call logs its duration so
 *   operators can identify latency bottlenecks.
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
const { Pool } = require("undici");

// ── Timeout Defaults ──────────────────────────────────────────────────────────

/** Default timeout for SSO login requests (15 s). */
const DEFAULT_SSO_TIMEOUT_MS = 15_000;

/** Default timeout for Agent `/jobs/sync` calls (5 min). */
const DEFAULT_AGENT_TIMEOUT_MS = 300_000;

/** Default timeout for Document Creator `/jobs/sync` calls (2 min). */
const DEFAULT_DOCUMENT_TIMEOUT_MS = 120_000;

/**
 * Parse a timeout config value to an integer, returning the default if
 * the value is missing or invalid.
 *
 * @param {string|undefined} raw      - The raw config string.
 * @param {number}           fallback - Default timeout in ms.
 * @returns {number} The parsed timeout in milliseconds.
 * @private
 */
function _parseTimeout(raw, fallback) {
  const n = parseInt(raw, 10);
  return isNaN(n) || n <= 0 ? fallback : n;
}

// ── HTTP Connection Pool ──────────────────────────────────────────────────────

/**
 * Shared `undici.Pool` that keeps TCP/TLS connections alive to the
 * TotalAgility origin.  All `fetch()` calls in this module pass
 * `{ dispatcher: taPool }` so that the underlying socket is reused
 * across requests — eliminating the TLS handshake overhead that would
 * otherwise occur on every HTTP call.
 *
 * The pool is created lazily on first use (by `_getPool()`) because
 * `config.totalAgilityEndpoint` may not yet be populated at module-load
 * time (e.g. when env-cmd injects variables after require).
 *
 * @type {Pool|null}
 * @private
 */
let taPool = null;

/**
 * Return (and lazily create) the shared undici connection pool for the
 * TotalAgility origin.
 *
 * The pool is configured with:
 * - `connections: 10`  — up to 10 concurrent keep-alive sockets.
 * - `pipelining: 1`    — standard HTTP/1.1 pipelining (one request per
 *   socket at a time).
 * - `keepAliveTimeout` / `keepAliveMaxTimeout` — keep idle sockets open
 *   for up to 60 s before closing them, allowing rapid reuse for the
 *   next API call.
 *
 * @returns {Pool} The undici Pool instance.
 * @private
 */
function _getPool() {
  if (!taPool) {
    const origin = config.totalAgilityEndpoint;
    if (!origin) {
      throw new Error(
        "[taAgent] TOTALAGILITY_ENDPOINT is not configured — cannot create connection pool."
      );
    }
    // Extract just the origin (scheme + host + port) from the full URL.
    const url = new URL(origin);
    const poolOrigin = url.origin;

    taPool = new Pool(poolOrigin, {
      connections: 10,
      pipelining: 1,
      keepAliveTimeout: 60_000,
      keepAliveMaxTimeout: 60_000,
    });
    console.log(`[taAgent] Created undici Pool for origin: ${poolOrigin}`);
  }
  return taPool;
}

// ── Timed Fetch Helper ────────────────────────────────────────────────────────

/**
 * Execute a `fetch()` request with a timeout and the shared connection pool.
 *
 * Uses `AbortController` to enforce the timeout.  If the request exceeds
 * `timeoutMs`, the fetch is aborted and an `Error` is thrown with a
 * descriptive message.
 *
 * Also logs the elapsed time for every request to aid performance analysis.
 *
 * @param {string}        url       - The URL to fetch.
 * @param {RequestInit}   options   - Standard fetch options (method, headers, body).
 * @param {number}        timeoutMs - Maximum time to wait (milliseconds).
 * @param {string}        label     - Human-readable label for log messages.
 * @returns {Promise<Response>} The fetch Response.
 * @throws {Error} If the request times out or the fetch fails.
 * @private
 */
async function _timedFetch(url, options, timeoutMs, label) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);

  const start = Date.now();
  try {
    const response = await fetch(url, {
      ...options,
      signal: controller.signal,
      dispatcher: _getPool(),
    });
    const elapsed = Date.now() - start;
    console.log(`[taAgent] ${label} completed in ${elapsed} ms (status ${response.status})`);
    return response;
  } catch (err) {
    const elapsed = Date.now() - start;
    if (err.name === "AbortError") {
      throw new Error(
        `[taAgent] ${label} timed out after ${timeoutMs} ms (elapsed: ${elapsed} ms). URL: ${url}`
      );
    }
    console.error(`[taAgent] ${label} failed after ${elapsed} ms:`, err.message);
    throw err;
  } finally {
    clearTimeout(timer);
  }
}

// ── Session Management ────────────────────────────────────────────────────────

/**
 * Custom error thrown when TotalAgility returns a 403 with an "Invalid Session
 * ID" message, indicating the SSO session has expired or is otherwise invalid.
 *
 * This error is caught by the auth-aware wrapper functions
 * ({@link callRestServiceWithAuth}, {@link createTotalAgilityDocumentWithAuth})
 * which clear the stale session, request a fresh one, and retry the call once.
 *
 * @extends Error
 */
class InvalidSessionError extends Error {
  /**
   * @param {string} message - The TotalAgility error message.
   */
  constructor(message) {
    super(message);
    this.name = "InvalidSessionError";
  }
}

/**
 * In-memory cache of TotalAgility SSO session keys, keyed by user identifier
 * (email address or test username).
 *
 * Sessions are reused across API calls until they expire or are explicitly
 * invalidated by a 403 response. This avoids the overhead of requesting a
 * new SSO session for every single API call.
 *
 * @type {Map<string, string>}
 * @private
 */
const sessionCache = new Map();

/**
 * In-flight SSO login promises, keyed by user identifier.
 *
 * When multiple concurrent requests trigger a session refresh for the same
 * user, only one SSO HTTP call is made; subsequent callers await the same
 * promise.  This prevents a thundering-herd of duplicate SSO logins and
 * reduces latency for concurrent requests.
 *
 * @type {Map<string, Promise<string>>}
 * @private
 */
const ssoInflight = new Map();

/**
 * Resolve a consistent user identifier from the current bot turn context.
 *
 * - In test mode (`TOTALAGILITY_USE_TEST_USER=true`): returns the configured
 *   test username.
 * - In production: resolves the Teams user's email address via
 *   `TeamsInfo.getMember()`.
 *
 * @param {import("botbuilder").TurnContext} context - The current bot turn context.
 * @returns {Promise<string|null>} The user identifier, or `null` if unresolvable.
 * @private
 */
async function _resolveUserId(context) {
  if (config.totalAgilityUseTestUser === "true") {
    return config.totalAgilityTestUserName;
  }
  const userInfo = await getCurrentUserIdAndEmail(context);
  return userInfo ? userInfo.email : null;
}

/**
 * Obtain a valid TotalAgility SSO session key for the current user.
 *
 * Returns a cached session if one exists (and `forceRefresh` is false),
 * otherwise requests a new session via {@link taSSOLogin} and stores it
 * in the cache.
 *
 * **Deduplication (v1.9):** if an SSO login is already in flight for this
 * user, the existing promise is returned instead of starting a second
 * concurrent login.  This prevents a thundering-herd when multiple API
 * calls fail with 403 simultaneously.
 *
 * @param {import("botbuilder").TurnContext} context - The current bot turn context.
 * @param {boolean} [forceRefresh=false] - When `true`, bypasses the cache and
 *   always requests a new session from TotalAgility.
 * @returns {Promise<string>} The TotalAgility session key.
 * @throws {Error} If user identity cannot be resolved or SSO login fails.
 * @private
 */
async function _getSession(context, forceRefresh = false) {
  const userId = await _resolveUserId(context);
  if (!userId) {
    throw new Error("Unable to resolve user identity for TotalAgility SSO.");
  }

  // Return cached session if available and not forcing a refresh.
  if (!forceRefresh && sessionCache.has(userId)) {
    console.log(`[taAgent] Using cached session for user: ${userId}`);
    return sessionCache.get(userId);
  }

  // Deduplicate: if an SSO login is already in flight for this user,
  // return the existing promise instead of starting a second one.
  if (ssoInflight.has(userId)) {
    console.log(`[taAgent] Waiting for in-flight SSO login for user: ${userId}`);
    return ssoInflight.get(userId);
  }

  console.log(`[taAgent] Requesting new SSO session for user: ${userId}`);

  // Create the login promise and register it for deduplication.
  const loginPromise = taSSOLogin(context)
    .then((sessionKey) => {
      sessionCache.set(userId, sessionKey);
      return sessionKey;
    })
    .finally(() => {
      // Always clean up the in-flight tracker, regardless of success/failure.
      ssoInflight.delete(userId);
    });

  ssoInflight.set(userId, loginPromise);
  return loginPromise;
}

/**
 * Remove the cached session for the current user, forcing a fresh SSO login
 * on the next API call.
 *
 * Called automatically when a 403 "Invalid Session ID" response is received.
 *
 * @param {import("botbuilder").TurnContext} context - The current bot turn context.
 * @returns {Promise<void>}
 * @private
 */
async function _clearSession(context) {
  const userId = await _resolveUserId(context);
  if (userId && sessionCache.has(userId)) {
    console.log(`[taAgent] Clearing cached session for user: ${userId}`);
    sessionCache.delete(userId);
  }
}

/**
 * Inspect a `fetch` Response for a TotalAgility 403 "Invalid Session ID" error.
 *
 * If the response status is 403 and the JSON body contains an `ErrorMessage`
 * matching "Invalid Session ID", an {@link InvalidSessionError} is thrown so
 * callers can trigger a session refresh and retry.
 *
 * For non-403 responses this function is a no-op.  For 403 responses that are
 * **not** invalid-session errors, a generic `Error` is thrown.
 *
 * @param {Response} response - The `fetch` Response object.
 * @param {string}   url      - The request URL (for logging).
 * @throws {InvalidSessionError} If the response is a 403 with an invalid session message.
 * @throws {Error} If the response is a 403 for any other reason.
 * @private
 */
async function _checkForInvalidSession(response, url) {
  if (response.status === 403) {
    let body;
    try {
      body = await response.json();
    } catch (_) {
      // Could not parse the response body — treat as a generic 403.
    }

    if (
      body &&
      body.ErrorMessage &&
      String(body.ErrorMessage).includes("Invalid Session ID")
    ) {
      console.warn("[taAgent] Invalid session detected:", body.ErrorMessage);
      throw new InvalidSessionError(body.ErrorMessage);
    }

    // 403 but not an invalid session — throw a generic HTTP error.
    const detail = body && body.ErrorMessage ? body.ErrorMessage : "Forbidden";
    throw new Error(`HTTP 403: ${detail}. URL: ${url}`);
  }
}

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
 * **Note:** This low-level function does not manage sessions. Use
 * {@link createTotalAgilityDocumentWithAuth} for automatic session management.
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
 * @throws {InvalidSessionError} If TotalAgility returns a 403 "Invalid Session ID".
 */
async function createTotalAgilityDocument(base64String, mimeType, sessionKey, fileName) {
  console.log("[taAgent] createTotalAgilityDocument() called for:", fileName);

  if (!base64String) return "";

  const url = config.totalAgilityEndpoint + "/jobs/sync";
  const timeoutMs = _parseTimeout(
    config.totalAgilityDocumentTimeoutMs,
    DEFAULT_DOCUMENT_TIMEOUT_MS
  );

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
    const response = await _timedFetch(
      url,
      { method: "POST", headers, body: JSON.stringify(payload) },
      timeoutMs,
      `Document Creator (${fileName})`
    );

    // Check for expired/invalid session (throws InvalidSessionError).
    await _checkForInvalidSession(response, url);

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
    // Re-throw session errors so the auth wrapper can handle them.
    if (error instanceof InvalidSessionError) throw error;

    console.error("[taAgent] Document Creator error:", error);
    return "";
  }
}

/**
 * Pre-create a TotalAgility Document **with automatic session management**.
 *
 * This is the recommended entry point for document creation. It:
 * 1. Obtains a session key (from cache or via SSO).
 * 2. Calls {@link createTotalAgilityDocument}.
 * 3. On a 403 "Invalid Session ID" error, clears the stale session, requests
 *    a fresh one, and retries the call exactly **once**.
 * 4. Returns an empty string on any authentication failure (allowing the
 *    caller in `teamsBot.js` to fall back to inline base64 mode).
 *
 * @param {string} base64String - Base64-encoded file content.
 * @param {string} mimeType     - MIME type of the file.
 * @param {import("botbuilder").TurnContext} context - The current bot turn context (used for session management).
 * @param {string} fileName     - Original filename.
 * @returns {Promise<string>} The TotalAgility Document ID, or empty string on failure.
 */
async function createTotalAgilityDocumentWithAuth(base64String, mimeType, context, fileName) {
  if (!base64String) return "";

  // Step 1: Obtain a session key (cached or new).
  let sessionKey;
  try {
    sessionKey = await _getSession(context);
  } catch (ssoErr) {
    console.error("[taAgent] SSO login failed for document creation:", ssoErr.message);
    return "";
  }

  // Step 2: Attempt the document creation.
  try {
    return await createTotalAgilityDocument(base64String, mimeType, sessionKey, fileName);
  } catch (err) {
    if (err instanceof InvalidSessionError) {
      // Step 3: Session expired — clear and retry once.
      console.log("[taAgent] Session expired during document creation, retrying...");
      await _clearSession(context);

      try {
        sessionKey = await _getSession(context, true);
      } catch (ssoErr) {
        console.error("[taAgent] SSO re-login failed for document creation:", ssoErr.message);
        return "";
      }

      try {
        return await createTotalAgilityDocument(base64String, mimeType, sessionKey, fileName);
      } catch (retryErr) {
        if (retryErr instanceof InvalidSessionError) {
          console.error(
            "[taAgent] Session still invalid for document creation after refresh — giving up."
          );
        } else {
          console.error("[taAgent] Document creation error on retry:", retryErr);
        }
        return "";
      }
    }

    // Non-session error (createTotalAgilityDocument already returns "" for
    // most errors, so this path is a safety net).
    console.error("[taAgent] Unexpected error in createTotalAgilityDocumentWithAuth:", err);
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
 * **Note:** This low-level function does not manage sessions. Use
 * {@link callRestServiceWithAuth} for automatic session management.
 *
 * @param {string} prompt_text   - The full conversation history rendered as Markdown.
 * @param {string} sessionKey    - TotalAgility SSO session ID (used as `Authorization` header).
 * @param {Object|null} [documentInfo=null] - Document attachment for this turn, or null/undefined if none.
 * @param {string} [documentInfo.documentId]  - TotalAgility Document ID (preload mode).
 * @param {string} [documentInfo.base64String] - Base64-encoded file content (inline mode).
 * @param {string} [documentInfo.mimeType]     - MIME type of the file (inline mode).
 * @param {string} [documentInfo.fileName]     - Original filename (inline mode).
 * @returns {Promise<string>} The agent's text response, or an error message string.
 * @throws {InvalidSessionError} If TotalAgility returns a 403 "Invalid Session ID".
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
  const timeoutMs = _parseTimeout(
    config.totalAgilityAgentTimeoutMs,
    DEFAULT_AGENT_TIMEOUT_MS
  );
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
    const response = await _timedFetch(
      url,
      { method: "POST", headers, body: JSON.stringify(payload) },
      timeoutMs,
      "Agent /jobs/sync"
    );

    // Check for expired/invalid session (throws InvalidSessionError).
    await _checkForInvalidSession(response, url);

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
    // Re-throw session errors so the auth wrapper can handle retry logic.
    if (error instanceof InvalidSessionError) throw error;

    console.error("Error: ", error);
    return_response =
      "Error: " + error + "\nPayload: " + JSON.stringify(payload);
  }

  return return_response;
}

/**
 * Call the TotalAgility Agent **with automatic session management**.
 *
 * This is the **primary entry point** for sending prompts to the Agent from
 * `teamsBot.js`.  It handles the full session lifecycle:
 *
 * 1. **Obtain a session** — retrieves a cached SSO session key for the
 *    current user, or requests a new one via {@link taSSOLogin} if none
 *    exists.
 * 2. **Call the Agent** — delegates to {@link callRestService}.
 * 3. **Retry on session expiry** — if the call returns a 403 "Invalid
 *    Session ID" error, the stale session is cleared, a fresh session is
 *    requested, and the call is retried exactly **once**.
 * 4. **Loop prevention** — if the retry also returns a 403, a user-friendly
 *    error message is returned instead of retrying again, preventing an
 *    infinite loop.
 * 5. **Graceful SSO failure** — if the SSO login itself fails (either on
 *    initial login or during the retry), a user-friendly error message is
 *    returned.
 *
 * @param {import("botbuilder").TurnContext} context - The current bot turn context.
 * @param {string} prompt_text   - The full conversation history rendered as Markdown.
 * @param {Object|null} [documentInfo=null] - Document attachment for this turn.
 * @param {string} [documentInfo.documentId]  - TotalAgility Document ID (preload mode).
 * @param {string} [documentInfo.base64String] - Base64-encoded file content (inline mode).
 * @param {string} [documentInfo.mimeType]     - MIME type of the file (inline mode).
 * @param {string} [documentInfo.fileName]     - Original filename (inline mode).
 * @returns {Promise<string>} The agent's text response, or a user-facing error message.
 */
async function callRestServiceWithAuth(context, prompt_text, documentInfo) {
  let sessionKey;

  // Step 1: Obtain a session key (cached or new).
  try {
    sessionKey = await _getSession(context);
  } catch (ssoErr) {
    console.error("[taAgent] SSO login failed:", ssoErr.message);
    return (
      "⚠️ Unable to sign into TotalAgility. Please try again in a moment.\n\n" +
      `Error: ${ssoErr.message}`
    );
  }

  // Step 2: Call the Agent API.
  try {
    return await callRestService(prompt_text, sessionKey, documentInfo);
  } catch (err) {
    if (err instanceof InvalidSessionError) {
      // Step 3: Session expired — clear the stale session and retry once.
      console.log("[taAgent] Session expired, requesting new session and retrying...");
      await _clearSession(context);

      let freshSessionKey;
      try {
        freshSessionKey = await _getSession(context, true);
      } catch (ssoErr) {
        console.error("[taAgent] SSO re-login failed:", ssoErr.message);
        return (
          "⚠️ Unable to re-authenticate with TotalAgility after session expiry.\n\n" +
          `Error: ${ssoErr.message}`
        );
      }

      // Step 4: Retry with the fresh session.
      try {
        return await callRestService(prompt_text, freshSessionKey, documentInfo);
      } catch (retryErr) {
        if (retryErr instanceof InvalidSessionError) {
          // Loop prevention: two consecutive 403s → give up gracefully.
          console.error("[taAgent] Session still invalid after refresh — giving up.");
          return (
            "⚠️ Unable to establish a valid TotalAgility session. " +
            "Your session was refreshed but remains invalid. " +
            "Please contact your administrator."
          );
        }
        // Non-session error on retry.
        console.error("[taAgent] Error on retry:", retryErr);
        return `⚠️ An error occurred while calling the TotalAgility Agent.\n\nError: ${retryErr.message}`;
      }
    }

    // Non-session error on first attempt. (callRestService already catches
    // most errors and returns strings, so this path is a safety net.)
    console.error("[taAgent] Unexpected error in callRestServiceWithAuth:", err);
    return `⚠️ An error occurred while calling the TotalAgility Agent.\n\nError: ${err.message}`;
  }
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
 * **Note:** Callers should prefer {@link callRestServiceWithAuth} and
 * {@link createTotalAgilityDocumentWithAuth} which manage sessions
 * automatically — only use `taSSOLogin` directly if you need low-level
 * session control.
 *
 * @param {import("botbuilder").TurnContext} context - The current bot turn context.
 * @returns {Promise<string>} The TotalAgility session ID.
 * @throws {Error} If the SSO HTTP request fails.
 */
async function taSSOLogin(context) {
  try {
    const ssoUrl =
      config.totalAgilityEndpoint + "/users/sessions/single-sign-on";
    const timeoutMs = _parseTimeout(
      config.totalAgilitySsoTimeoutMs,
      DEFAULT_SSO_TIMEOUT_MS
    );

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

    const ssoResponse = await _timedFetch(
      ssoUrl,
      { method: "POST", headers: ssoHeaders, body: JSON.stringify(ssoPayload) },
      timeoutMs,
      "SSO login"
    );

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
    throw error; // Bubble up — caught by _getSession or callRestServiceWithAuth.
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
  callRestServiceWithAuth,
  createTotalAgilityDocument,
  createTotalAgilityDocumentWithAuth,
  taSSOLogin,
};
