/*************************************************************************/
/*  Conversation reference store backed by Azure Table Storage.         */
/*  Falls back to in-memory when no connection string is configured.    */
/*                                                                       */
/*  Used for proactive (push) notifications — the official Microsoft    */
/*  pattern for sending messages to specific users from external apps.  */
/*  See: https://learn.microsoft.com/en-us/microsoftteams/platform/     */
/*       bots/how-to/conversations/send-proactive-messages              */
/*************************************************************************/

// NOTE: @azure/data-tables is loaded lazily inside init() so that the app
// can start even if the package has issues or AZURE_STORAGE_CONNECTION_STRING
// is not configured.  This avoids a crash on Azure App Service when the
// SDK's transitive dependencies are incompatible with the host Node version.

const config = require("./config");

const TABLE_NAME = "ConversationReferences";
const PARTITION_KEY = "TeamsBot"; // single-partition for simplicity

let tableClient = null;
let useTable = false;

// In-memory fallback (used when AZURE_STORAGE_CONNECTION_STRING is absent)
const memoryStore = new Map();

/**
 * Initialise the Azure Table Storage client and ensure the table exists.
 * Call once at startup.  Safe to call multiple times — will no-op after first.
 */
async function init() {
  if (tableClient) return; // already initialised

  const connStr = config.azureStorageConnectionString;
  if (!connStr) {
    console.warn(
      "[ConversationStore] No AZURE_STORAGE_CONNECTION_STRING configured — " +
        "falling back to in-memory store.  References will be lost on restart."
    );
    return;
  }

  try {
    // Lazy-load the Azure SDK so the app can start even if the package
    // is missing or has compatibility issues.
    const { TableClient } = require("@azure/data-tables");

    tableClient = TableClient.fromConnectionString(connStr, TABLE_NAME, {
      allowInsecureConnection: connStr.includes("UseDevelopmentStorage=true"),
    });
    await tableClient.createTable(); // no-op if already exists
    useTable = true;
    console.log("[ConversationStore] Azure Table Storage initialised (table: %s).", TABLE_NAME);
  } catch (err) {
    // If the table already exists the SDK throws a 409 — that's fine.
    if (err.statusCode === 409) {
      useTable = true;
      console.log("[ConversationStore] Azure Table Storage table already exists — ready.");
    } else {
      console.error("[ConversationStore] Failed to initialise Azure Table Storage:", err.message);
      console.warn("[ConversationStore] Falling back to in-memory store.");
      tableClient = null;
    }
  }
}

/**
 * Save (or overwrite) a conversation reference for a user.
 *
 * @param {string} userKey   - Unique user identifier (email or Teams name).
 * @param {Partial<import("botbuilder").ConversationReference>} ref
 */
async function save(userKey, ref) {
  if (!userKey || !ref) return;
  const key = userKey.toLowerCase();

  if (useTable && tableClient) {
    const entity = {
      partitionKey: PARTITION_KEY,
      rowKey: encodeURIComponent(key), // row keys must be URL-safe
      userKey: key,
      referenceJson: JSON.stringify(ref),
    };
    try {
      await tableClient.upsertEntity(entity, "Replace");
    } catch (err) {
      console.error("[ConversationStore] Table upsert failed for %s:", key, err.message);
      // Also store in memory as a safety net
      memoryStore.set(key, ref);
    }
  } else {
    memoryStore.set(key, ref);
  }
}

/**
 * Retrieve a previously-stored conversation reference.
 *
 * @param {string} userKey
 * @returns {Promise<Partial<import("botbuilder").ConversationReference> | undefined>}
 */
async function get(userKey) {
  if (!userKey) return undefined;
  const key = userKey.toLowerCase();

  if (useTable && tableClient) {
    try {
      const entity = await tableClient.getEntity(PARTITION_KEY, encodeURIComponent(key));
      return JSON.parse(entity.referenceJson);
    } catch (err) {
      // 404 = not found — expected for unknown users
      if (err.statusCode !== 404) {
        console.error("[ConversationStore] Table get failed for %s:", key, err.message);
      }
      // Fall through to memory
      return memoryStore.get(key);
    }
  }

  return memoryStore.get(key);
}

/**
 * Return an array of all known user keys.
 *
 * @returns {Promise<string[]>}
 */
async function listUsers() {
  if (useTable && tableClient) {
    try {
      const users = [];
      const entities = tableClient.listEntities({
        queryOptions: { select: ["userKey"] },
      });
      for await (const entity of entities) {
        users.push(entity.userKey);
      }
      return users;
    } catch (err) {
      console.error("[ConversationStore] listUsers table query failed:", err.message);
      // Fall through to memory store
    }
  }

  return Array.from(memoryStore.keys());
}

/**
 * Return a diagnostic snapshot of all stored references.
 *
 * @returns {Promise<{ userKey: string, conversationId: string | null, userName: string | null }[]>}
 */
async function listAll() {
  if (useTable && tableClient) {
    try {
      const result = [];
      const entities = tableClient.listEntities();
      for await (const entity of entities) {
        const ref = JSON.parse(entity.referenceJson);
        result.push({
          userKey: entity.userKey,
          conversationId: ref.conversation ? ref.conversation.id : null,
          userName: ref.user ? ref.user.name : null,
        });
      }
      return result;
    } catch (err) {
      console.error("[ConversationStore] listAll table query failed:", err.message);
      // Fall through to memory store
    }
  }

  const result = [];
  memoryStore.forEach((ref, key) => {
    result.push({
      userKey: key,
      conversationId: ref.conversation ? ref.conversation.id : null,
      userName: ref.user ? ref.user.name : null,
    });
  });
  return result;
}

module.exports = { init, save, get, listUsers, listAll };
