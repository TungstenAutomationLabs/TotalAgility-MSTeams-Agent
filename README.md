
# Teams Chat Client for TotalAgility Agents

This repository is a lightweight Microsoft Teams **Chat Client** (using the Microsoft Bot Framework) that proxies chat interactions to TotalAgility **Chat Agents** (or custom LLM-backed Agents) via REST APIs. It displays Chat Agent responses in Teams while leaving the automation, data access, and LLM logic inside TotalAgility.

![Example screenshot of an interaction with a TotalAgility AI Agent](img/TotalAgility-Teams-App-Screenshot.png)

---

## Table of Contents

- [Key Capabilities](#key-capabilities)
- [Why Use This Sample](#why-use-this-sample)
- [Architecture Summary](#architecture-summary)
- [Supported Channels](#supported-channels)
- [Getting Started](#getting-started)
  - [Prerequisites](#prerequisites)
  - [Step-by-Step Deployment](#step-by-step-deployment)
- [Project Structure](#project-structure)
  - [Important Files to Edit](#important-files-to-edit)
  - [Editing Tips](#editing-tips)
- [Configuration](#configuration)
  - [Environment Variables](#environment-variables)
  - [Environment Variable Reference](#environment-variable-reference)
  - [Behaviour Notes](#behaviour-notes)
  - [Built-in Chat Commands](#built-in-chat-commands)
- [Intent Router Pattern](#intent-router-pattern)
- [Prompt Examples — Teams Message Formatting](#prompt-examples--teams-message-formatting)
- [Proactive Notifications API](#proactive-notifications-api)
  - [Finding Your Endpoint URL](#finding-your-endpoint-url)
  - [POST /api/notifications](#post-apinotifications)
  - [GET /api/conversations](#get-apiconversations)
  - [How It Works](#how-it-works)
- [Document Preloading (Recommended for Production)](#document-preloading-recommended-for-production)
- [Deployment and Infrastructure](#deployment-and-infrastructure)
- [Version History](#version-history)
- [Microsoft 365 Agents Toolkit — Resources](#microsoft-365-agents-toolkit--resources)
- [Example Resources](#example-resources)
- [Related Projects](#related-projects)

---

## Key Capabilities

- Proxy Teams chat to a TotalAgility Chat Agent (or Custom LLM) using REST/OpenAPI
- Preserve and pass conversation history to the Chat Agent to maintain context
- Support for sending documents (file upload) to the Chat Agent (see Version notes)
- Send messages to users, for example to notify of state changes in your Agent or Case
- Minimal Chat Client logic so the Chat Agent process in TotalAgility can own orchestration, knowledge access and calling other sub-agents

## Why Use This Sample

It demonstrates how Teams can be used as a Chat Client for complex automation and LLM-driven processes implemented in TotalAgility (workflows, case management, RPA, IDP, knowledge base lookups, document generation, eSigning, etc.).

## Architecture Summary

- The Teams Chat Client acts as a thin proxy: it accepts user messages, forwards them (with conversation history) to a TotalAgility Chat Agent endpoint, then renders the Chat Agent's response in Teams.
- The recommended pattern is an "Intent Router" Chat Agent in TotalAgility that receives the prompt, evaluates intent via an LLM step, and routes to specific sub-agents/processes (each with the same Chat Agent interface).

## Supported Channels

- Microsoft Teams (primary)
- Any Bot Framework channel (Webchat, Facebook Messenger, WhatsApp, Alexa, etc.) with minimal changes.

---

## Getting Started

### Prerequisites

Before you begin, ensure you have the following installed and configured:

| # | Prerequisite | Description | Link |
|---|-------------|-------------|------|
| 1 | **Visual Studio Code** | The recommended IDE. The project includes VS Code workspace tasks and launch configurations. | [Download VS Code](https://code.visualstudio.com/) |
| 2 | **Microsoft 365 Agents Toolkit** (VS Code extension) | Formerly "Teams Toolkit". Provides project scaffolding, local tunnelling, provisioning, and deployment commands for Teams apps. Install from the VS Code Extensions marketplace. | [Install Agents Toolkit](https://learn.microsoft.com/en-us/microsoftteams/platform/toolkit/install-agents-toolkit) |
| 3 | **Microsoft 365 developer sandbox** | A free M365 developer tenant for testing Teams apps without affecting a production environment. Sign up via the Microsoft 365 Developer Program. | [M365 Developer Program](https://developer.microsoft.com/en-us/microsoft-365/dev-program) |
| 4 | **Azure subscription (sandbox or dev)** | Required for provisioning the Azure Bot Service and hosting the Chat Client in Azure App Service. A free Azure trial or Visual Studio subscriber credits work well for development. | [Azure Free Account](https://azure.microsoft.com/en-gb/free/) |
| 5 | **Node.js (LTS)** | Required to run the Chat Client locally and install dependencies via npm. Version 18.x or 20.x LTS recommended. | [Download Node.js](https://nodejs.org/) |
| 6 | **Git** | Version control — needed to clone this repository. | [Download Git](https://git-scm.com/) |
| 7 | *(Optional)* **Azurite** | Local Azure Storage emulator — useful for testing Azure Table Storage (conversation references) without an Azure account. Included as a VS Code extension or installable via npm. | [Azurite on npm](https://www.npmjs.com/package/azurite) |

### Step-by-Step Deployment

#### 1. Clone the repository

```bash
git clone <repository-url>
cd TotalAgility-Agent
```

#### 2. Install dependencies

```bash
npm install
```

#### 3. Install the Microsoft 365 Agents Toolkit extension

Open VS Code, go to the Extensions panel (`Ctrl+Shift+X`), and search for **"Microsoft 365 Agents Toolkit"** (previously called "Teams Toolkit"). Install the extension and sign in with both your **M365 developer account** and your **Azure account** when prompted.

#### 4. Configure environment variables

Copy the environment template files in `env/` and fill in your values:

- `env/.env.local` — for local development
- `env/.env.local.user` — for secrets (API keys, tokens) during local dev
- `env/.env.dev` — for Azure deployment (populated automatically during provisioning)
- `env/.env.dev.user` — for secrets used in Azure deployment

At a minimum, set:
```
TOTALAGILITY_ENDPOINT=https://<your_tenant>.dev.kofaxcloud.com/services/sdk/v1
SECRET_TOTALAGILITY_API_KEY=<your-api-key>
TOTALAGILITY_AGENT_NAME=<your-agent-process-name>
TOTALAGILITY_AGENT_ID=<your-agent-process-id>
```

Refer to the [Environment Variables](#environment-variables) section below for the full list.

#### 5. Run locally with the Agents Toolkit

Use one of the following methods:

**Option A — VS Code tasks (recommended):**
1. Open the Command Palette (`Ctrl+Shift+P`) and select **"Microsoft 365 Agents Toolkit: Preview Your Teams App"**.
2. Choose **"Local"** as the environment.
3. The toolkit will automatically start a local dev tunnel, provision a temporary bot registration, and launch the app.

**Option B — npm scripts:**
```bash
npm run dev:teamsfx           # start locally for Teams development
npm run dev:teamsfx:testtool  # start using the Teams App Test Tool
```

#### 6. Test in Microsoft Teams

1. Open Microsoft Teams (desktop or web) using your **M365 developer sandbox** account.
2. The Agents Toolkit will sideload the app automatically when previewing locally.
3. Find the app in the Teams chat and send a message to verify the connection to your TotalAgility Chat Agent.

#### 7. Provision Azure resources

When you are ready to deploy to Azure:

1. Open the Command Palette (`Ctrl+Shift+P`) and select **"Microsoft 365 Agents Toolkit: Provision"**.
2. Select your Azure subscription and choose a resource group (or create a new one).
3. The toolkit runs the Bicep templates in `infra/` to create the Azure Bot Service, App Service, and App Service Plan.
4. After provisioning, check `env/.env.dev` — the toolkit writes `BOT_ID`, `BOT_DOMAIN`, and other values here automatically.

#### 8. Deploy to Azure

1. Open the Command Palette and select **"Microsoft 365 Agents Toolkit: Deploy"**.
2. The toolkit packages and deploys the Chat Client to the provisioned Azure App Service.
3. Update `env/.env.dev.user` with your production secrets (`SECRET_TOTALAGILITY_API_KEY`, `SECRET_NOTIFICATIONS_BEARER_TOKEN`, `SECRET_AZURE_STORAGE_CONNECTION_STRING`).

#### 9. Publish the app to your organisation

1. Open the Command Palette and select **"Microsoft 365 Agents Toolkit: Publish"**.
2. This submits the app package to your M365 tenant's app catalogue for admin approval.
3. Once approved, users in your organisation can install the Chat Client from the Teams app store.

> **Tip:** For a quick test without publishing, use **"Microsoft 365 Agents Toolkit: Zip Teams App Package"** and sideload the generated `.zip` file manually in Teams.

---

## Project Structure

- `appPackage/` — Teams app manifest and package. Example prompts ("View prompts" above the Send button in the Teams app) and app configuration live in `appPackage/manifest.json`.
- `src/` — core application code
	- `src/index.js` — app entry point and server bootstrap
	- `src/config.js` — configuration loader and environment handling
	- `src/taAgent.js` — main TotalAgility API integration and the primary place to change how the app calls your Agent (seed usage, request shape, headers, error handling)
	- `src/teamsBot.js` — Teams/Microsoft Bot framework adapter and conversation turn handling
	- `src/conversationStore.js` — Azure Table Storage–backed persistence for conversation references (proactive messaging)
	- `src/utils.js` — helper functions including loading/typing feedback messages (customize UI text here)
- `prompt_examples/` — example prompts for use with your TotalAgility Chat Agents
	- `prompt_examples/example_teams_formatting_prompt.md` — a ready-to-use prompt that instructs an AI to format Teams messages using Tungsten Automation branding (brand colours, logos, HTML formatting patterns that are compatible with the Teams message renderer). Copy or adapt this prompt into your TotalAgility Chat Agent's LLM system prompt to produce consistently branded, accessible notifications and replies in Teams. See [Prompt Examples — Teams Message Formatting](#prompt-examples--teams-message-formatting) below.
- `env/` — environment template files — update these to match your tenant, keys and Agent details
- `infra/` — infrastructure deployment scripts (Azure Bicep templates used for provisioning, if desired)
- `devTools/` — developer utilities (Teams App Tester, etc.)

### Important Files to Edit

| Task | File(s) |
|------|---------|
| Change the Agent integration or modify request details | `src/taAgent.js` |
| Adjust loading/typing UI and helper utilities | `src/utils.js` |
| Teams-specific behaviour or message formatting | `src/teamsBot.js` and `src/index.js` |
| Environment-specific secrets | `env/` templates and local environment variables |

### Editing Tips

- To change how the app calls TotalAgility (payload, headers, error handling), update `src/taAgent.js`.
- To change user-visible loading/typing messages, update `src/utils.js`.
- For Teams-specific behaviour or message formatting, update `src/teamsBot.js` and `src/index.js`.
- Keep environment-specific secrets out of source control. Use the `env/` templates and local environment variables when running locally.

---

## Configuration

### Environment Variables

Place your environment values in the files under `env/` (or in your system environment). The sample adds the following keys:

```
TOTALAGILITY_ENDPOINT=
TOTALAGILITY_API_KEY=
TOTALAGILITY_AGENT_NAME=
TOTALAGILITY_AGENT_ID=
TOTALAGILITY_TEST_USERNAME=
TOTALAGILITY_USE_TEST_USER=
CONVERSATION_HISTORY_MAX_ENTRIES=
NOTIFICATIONS_BEARER_TOKEN=
AZURE_STORAGE_CONNECTION_STRING=
PRELOAD_DOCUMENTS_AS_TOTALAGILITY_DOCS=
TOTALAGILITY_DOCUMENT_CREATOR_PROCESS_ID=
TOTALAGILITY_DOCUMENT_CREATOR_PROCESS_NAME=
TOTALAGILITY_DOCUMENT_TYPE_ID=
TOTALAGILITY_DOCUMENT_FILENAME_FIELD_ID=
```

> **Security note — `SECRET_` prefix convention:**
> Teams Toolkit automatically masks variables whose names start with `SECRET_` in
> build and deploy logs.  Sensitive values (API keys, tokens, connection strings)
> should use `SECRET_` prefixed variable names in `env/.env.*.user` files
> (e.g. `SECRET_TOTALAGILITY_API_KEY`, `SECRET_NOTIFICATIONS_BEARER_TOKEN`,
> `SECRET_AZURE_STORAGE_CONNECTION_STRING`).  The YAML files map these to the
> non-prefixed names the application code expects.

### Environment Variable Reference

- `TOTALAGILITY_ENDPOINT` — base TotalAgility REST/OpenAPI endpoint, e.g. `https://{{your_tenant}}.dev.tungstencloud.com/services/sdk/v1`
- `TOTALAGILITY_API_KEY` — API key used to authenticate calls to TotalAgility
- `TOTALAGILITY_AGENT_NAME` — the process name of the Chat Agent in TotalAgility
- `TOTALAGILITY_AGENT_ID` — the process ID of the Chat Agent (often visible in the TotalAgility Designer edit URL)
- `TOTALAGILITY_TEST_USERNAME` & `TOTALAGILITY_USE_TEST_USER` — override SSO behaviour to force a test TA user (useful for development)
- `CONVERSATION_HISTORY_MAX_ENTRIES` — maximum number of messages to retain in the conversation history array sent to the Chat Agent. Defaults to `15` if not set or invalid. Higher values provide more context but increase payload size.

- `NOTIFICATIONS_BEARER_TOKEN` — a secret token that 3rd-party callers must present as a `Bearer` token when calling the notification endpoints. **Required** to enable `/api/notifications` and `/api/conversations`.
- `AZURE_STORAGE_CONNECTION_STRING` — connection string for an Azure Storage account used to persist conversation references in Azure Table Storage. When absent the app falls back to an in-memory store (references are lost on restart). For local development with Azurite use `UseDevelopmentStorage=true`.
- `PRELOAD_DOCUMENTS_AS_TOTALAGILITY_DOCS` — when `true`, uploaded files are first submitted to a dedicated "Document Creator" process in TotalAgility that stores the file and returns a Document ID.  That lightweight ID is then passed to the Chat Agent instead of the raw base64 content. Default: `false`. See [Document Preloading](#document-preloading-recommended-for-production) below.
- `TOTALAGILITY_DOCUMENT_CREATOR_PROCESS_ID` — the Process ID (GUID) of the TotalAgility Document Creator process. Required when `PRELOAD_DOCUMENTS_AS_TOTALAGILITY_DOCS=true`.
- `TOTALAGILITY_DOCUMENT_CREATOR_PROCESS_NAME` — the process name of the TotalAgility Document Creator process. Required when `PRELOAD_DOCUMENTS_AS_TOTALAGILITY_DOCS=true`.
- `TOTALAGILITY_DOCUMENT_TYPE_ID` — the Document Type ID (GUID) used when creating documents via the Document Creator process. Required when `PRELOAD_DOCUMENTS_AS_TOTALAGILITY_DOCS=true`. Default: `298D0A0CFE2342A4BB66E240E9E2967D` (the standard TotalAgility "Default Document Type" — this value is preset in all env templates and will work for most deployments).
- `TOTALAGILITY_DOCUMENT_FILENAME_FIELD_ID` — the RuntimeField ID (GUID) for the filename field on the document type. Required when `PRELOAD_DOCUMENTS_AS_TOTALAGILITY_DOCS=true`. Default: `1F8220766FAF42278F5CF8081DBF6D87` (preset in all env templates).

### Behaviour Notes

- The main API call is managed in `src/taAgent.js`. The Chat Client uses a hard-coded "seed" for consistent responses; remove or change this if you want nondeterministic LLM outputs.
- Loading messages can be configured in `src/utils.js`.
- Example prompts are available in `appPackage/manifest.json`.

### Built-in Chat Commands

| Command | Description |
|---------|-------------|
| `debug` | Prints all loaded configuration values and the current conversation history to the Teams chat (and console log). Sensitive values (API keys, passwords, tokens, connection strings) are masked. Useful for verifying that the correct environment files are being loaded at runtime and inspecting conversation context. |
| `clear conversation history` | Displays the current conversation history, then resets it. Also accepts: `clear history`, `clear`, `reset`, `clear conversation`. |

---

## Intent Router Pattern

A controlling "Intent Router" Chat Agent evaluates incoming prompts using an LLM step and maps them to available actions / sub-agents. The router provides a registry of available Chat Agents (with ProcessIDs) and returns JSON mapped to the TotalAgility data model describing which sub-process to invoke and which prompt to send.

Because all Chat Agents expose a common interface, the registry can include many agents; the router chooses the best match and may iterate multiple steps (search KB, call external APIs, gather documents) before returning a final response.

---

## Prompt Examples — Teams Message Formatting

The file [`prompt_examples/example_teams_formatting_prompt.md`](prompt_examples/example_teams_formatting_prompt.md) contains a ready-to-use system prompt that instructs an AI to format Teams messages using Tungsten Automation branding. It covers:

- **Brand colour palette** — Tungsten Automation colours mapped to inline `style=""` attributes (the only styling Teams supports)
- **Allowed & forbidden HTML tags** — a safe subset that Teams will render correctly
- **Formatting patterns** — section headings, tables, status indicators, alert blocks, links, images/logos
- **Complete message template** — a reusable skeleton for branded notifications and replies

Copy or adapt this prompt into your TotalAgility Chat Agent's LLM system prompt to produce consistently branded, accessible notifications and replies in Teams.

Below is an example of a formatted Teams message produced using this prompt:

![Example of a formatted Teams message using Tungsten Automation branding](img/Example-formatted-Teams-message.png)

---

## Proactive Notifications API

The Chat Client exposes two HTTP endpoints that allow external / 3rd-party applications to send messages directly into a user's Teams chat. This follows the [official Microsoft proactive messaging pattern](https://learn.microsoft.com/en-us/microsoftteams/platform/bots/how-to/conversations/send-proactive-messages).

**Prerequisites:**
1. Set the `NOTIFICATIONS_BEARER_TOKEN` environment variable to a strong secret.
2. Set the `AZURE_STORAGE_CONNECTION_STRING` environment variable for persistent storage (optional but recommended for production).
3. The target user must have interacted with the Chat Client at least once (or had it installed) so that their conversation reference is stored.

### Finding Your Endpoint URL

The notification endpoint URL depends on where the Chat Client is running:

| Environment | Base URL | How to find it |
|-------------|----------|----------------|
| **Local development** | `http://localhost:3978` | Default port configured in `src/index.js`. The Chat Client listens on port `3978` unless overridden by the `PORT` environment variable. |
| **Azure (deployed)** | `https://<BOT_DOMAIN>` | After running `teamsapp provision` and `teamsapp deploy`, the `BOT_DOMAIN` value is written to `env/.env.dev`. Open that file and look for a line like `BOT_DOMAIN=botb556a8.azurewebsites.net`. Your full URL is `https://` + that value. |

**Full endpoint URLs:**

| Endpoint | Local URL | Azure URL |
|----------|-----------|-----------|
| Send notification | `POST http://localhost:3978/api/notifications` | `POST https://<BOT_DOMAIN>/api/notifications` |
| List conversations | `GET http://localhost:3978/api/conversations` | `GET https://<BOT_DOMAIN>/api/conversations` |

**Where to find each configuration value:**

| Value | Where to find it |
|-------|-----------------|
| **Bot hostname** (`BOT_DOMAIN`) | `env/.env.dev` — populated after `teamsapp provision`. Example: `botb556a8.azurewebsites.net` |
| **Bearer token** | `SECRET_NOTIFICATIONS_BEARER_TOKEN` in your `env/.env.dev.user` (or `env/.env.local`, `env/.env.testtool` for local dev) |
| **User key** | The email address of any user who has messaged the Chat Client. Verify available users via `GET /api/conversations`. |

**Public access:** Azure App Service is publicly accessible by default over HTTPS (port 443). No additional networking configuration is needed — if the Bot Framework Channel Service can reach the Chat Client at `https://<BOT_DOMAIN>/api/messages`, then 3rd-party apps can reach `/api/notifications` at the same hostname.

> **Tip — restricting access:** If you want to limit which systems can call the notification endpoint (beyond bearer-token auth), configure **Azure App Service → Networking → Access Restrictions** in the Azure Portal to whitelist specific IP ranges.

### `POST /api/notifications`

Send a proactive message to a specific user.

**Headers:**
```
Authorization: Bearer <NOTIFICATIONS_BEARER_TOKEN>
Content-Type: application/json
```

**Request body:**
```json
{
  "userKey": "jane.doe@contoso.com",
  "message": "Your document has been processed successfully."
}
```

| Field | Type | Description |
|-------|------|-------------|
| `userKey` | string | The user's email address (or Teams display name) as registered when they last interacted with the Chat Client. Case-insensitive. |
| `message` | string | The message text to send to the user in their Teams chat. Supports Markdown. Max 4000 characters. |

**Responses:**

| Status | Meaning |
|--------|---------|
| `200`  | Message sent successfully. |
| `400`  | Missing `userKey` or `message` in request body. |
| `401`  | Invalid or missing bearer token. |
| `404`  | No conversation reference found for the given `userKey`. |
| `429`  | Rate limit exceeded — try again later. |
| `500`  | Internal server error. |
| `503`  | `NOTIFICATIONS_BEARER_TOKEN` is not configured. |

**Example (cURL):**
```bash
curl -X POST https://your-bot-host/api/notifications \
  -H "Authorization: Bearer YOUR_SECRET_TOKEN" \
  -H "Content-Type: application/json" \
  -d '{"userKey": "jane.doe@contoso.com", "message": "Job 12345 is complete."}'
```

**Example (JavaScript / Node.js):**
```javascript
const response = await fetch("https://your-bot-host/api/notifications", {
  method: "POST",
  headers: {
    "Authorization": "Bearer YOUR_SECRET_TOKEN",
    "Content-Type": "application/json",
  },
  body: JSON.stringify({
    userKey: "jane.doe@contoso.com",
    message: "**Document processed** ✅\nYour invoice #12345 has been approved.",
  }),
});

const result = await response.json();
console.log(result); // { status: "ok", userKey: "jane.doe@contoso.com", message: "..." }
```


**Example — calling from a TotalAgility process:**

In a TotalAgility process, use a **REST Service** activity to call the notification endpoint. Configure:
- **URL:** `https://your-bot-host/api/notifications`
- **Method:** `POST`
- **Headers:** `Authorization: Bearer YOUR_SECRET_TOKEN` and `Content-Type: application/json`
- **Body:** Map process variables to produce `{"userKey": "<email>", "message": "<notification text>"}`

This allows any TotalAgility workflow to push status updates directly into a user's Teams chat.

### `GET /api/conversations`

List all users with stored conversation references (useful for diagnostics and discovering valid `userKey` values).

**Headers:**
```
Authorization: Bearer <NOTIFICATIONS_BEARER_TOKEN>
```

**Response:**
```json
{
  "count": 2,
  "conversations": [
    {
      "userKey": "jane.doe@contoso.com",
      "conversationId": "a]b]c...",
      "userName": "Jane Doe"
    },
    {
      "userKey": "john.smith@contoso.com",
      "conversationId": "x]y]z...",
      "userName": "John Smith"
    }
  ]
}
```

### How It Works

1. Every time a user sends a message to the Chat Client (or the Chat Client is installed for a user), a `ConversationReference` is captured and stored — keyed by the user's email address (resolved via Teams APIs) or their display name as a fallback.
2. The conversation reference is persisted in Azure Table Storage (table: `ConversationReferences`) so it survives process restarts.
3. When a 3rd-party app calls `POST /api/notifications`, the Chat Client uses `adapter.continueConversationAsync()` with the stored reference to send the message into the user's existing personal chat.
4. The user receives the notification as a new message from the Chat Client in Teams — no user action required.

---

## Document Preloading (Recommended for Production)

When users upload files to the Chat Client, the default behaviour is to convert the file to a base64 string and pass it directly as an input variable (`DOCUMENT_CONTENT`) to the TotalAgility Chat Agent process (with `DOCUMENT` left empty).  While simple, this has a significant drawback: the entire base64 string is stored as a process variable in the TotalAgility database, which can be very large for multi-megabyte files.

**Document preloading** solves this by splitting the upload into two steps:

1. **Create the document** — the Chat Client calls a dedicated "Document Creator" process in TotalAgility (configured via `TOTALAGILITY_DOCUMENT_CREATOR_PROCESS_ID` / `TOTALAGILITY_DOCUMENT_CREATOR_PROCESS_NAME`).  The file is submitted as a document attachment in the `Documents` array (with `Base64Data`, `MimeType`, `DocumentTypeId`, and a `RuntimeFields` entry for the filename).  The response returns a top-level `DocumentId`.
2. **Call the Chat Agent** — the Chat Client calls the main Chat Agent process, passing the lightweight document reference via the `DOCUMENT` input variable (with `DOCUMENT_CONTENT`, `DOCUMENT_TYPE`, and `DOCUMENT_FILENAME` all empty) instead of the raw base64 string.  The Chat Agent can then retrieve the document from TotalAgility's document storage as needed.

**Benefits:**
- The document is stored once in TotalAgility's optimised document storage (not as a process variable).
- The Chat Agent process payload is much smaller, reducing database I/O and memory usage.
- Better suited for production deployments with large files or high throughput.

**How to enable:**
```
PRELOAD_DOCUMENTS_AS_TOTALAGILITY_DOCS=true
TOTALAGILITY_DOCUMENT_CREATOR_PROCESS_ID=<your-document-creator-process-id>
TOTALAGILITY_DOCUMENT_CREATOR_PROCESS_NAME=<your-document-creator-process-name>
TOTALAGILITY_DOCUMENT_TYPE_ID=298D0A0CFE2342A4BB66E240E9E2967D
TOTALAGILITY_DOCUMENT_FILENAME_FIELD_ID=1F8220766FAF42278F5CF8081DBF6D87
```

**Document Creator process requirements:**
The TotalAgility Document Creator process must:
1. Accept a document attachment in its `Documents` array.  The Chat Client submits the file with `Base64Data`, `MimeType`, `DocumentTypeId` (from `TOTALAGILITY_DOCUMENT_TYPE_ID`), and a `RuntimeFields` entry whose `Id` is `TOTALAGILITY_DOCUMENT_FILENAME_FIELD_ID` carrying the original filename.
2. Have `StoreFolderAndDocuments` enabled so the document is persisted in TotalAgility's document storage.
3. Return a top-level `DocumentId` in the `/jobs/sync` response (this is the standard TotalAgility behaviour when `StoreFolderAndDocuments=true` and `ReturnOnlySpecifiedDocuments=true`).

**Fallback:** If document preloading fails (e.g. the Document Creator process is unavailable), the Chat Client automatically falls back to sending the raw base64 string inline — so the user's request is not lost.

---

## Deployment and Infrastructure

The `infra/` folder contains Azure Bicep templates to help provision cloud resources if you want to deploy to Azure. See the [Step-by-Step Deployment](#step-by-step-deployment) section for a complete walkthrough.

---

## Version History

### Version 1.1
- Added the ability to upload files and send these to the Chat Agent for processing. This sample uses TotalAgility 25.2 where the Chat Agent interface accepts TotalAgility Documents (sent as base64 strings) to the Jobs sync API.

### Version 1.2
- Added settings to SSO a user into TotalAgility based on their email address from their Teams login.

*Note:* the code assumes the user's email address (from their MS Teams login) is their user ID in TotalAgility. To override this, specify a test user in the environment and set the `TOTALAGILITY_USE_TEST_USER` flag to `true`.

Environment variables for SSO testing:

```
TOTALAGILITY_TEST_USERNAME=my_ta_test_account@test.com
TOTALAGILITY_USE_TEST_USER=true
```

### Version 1.3
- Added **proactive notification endpoint** (`POST /api/notifications`) that allows 3rd-party systems (e.g. TotalAgility workflows, Power Automate, external APIs) to push messages into a specific user's Teams session.
- Added **conversation listing endpoint** (`GET /api/conversations`) to discover which users have active conversation references.
- Conversation references are persisted to **Azure Table Storage** for durability across restarts (falls back to in-memory when `AZURE_STORAGE_CONNECTION_STRING` is not set).
- Both endpoints are protected by bearer-token authentication via `NOTIFICATIONS_BEARER_TOKEN`.

### Version 1.4
- **Security hardening:** added `helmet` for HTTP security headers, `express-rate-limit` on notification endpoints, startup config validation, request body size limits.
- **Secret management:** sensitive env vars now use the `SECRET_` prefix convention so Teams Toolkit masks them in logs.
- Fixed `.gitignore` to prevent `.localConfigs` (which contains runtime secrets) from being committed.

### Version 1.5
- Added **document preloading** (`PRELOAD_DOCUMENTS_AS_TOTALAGILITY_DOCS`) — an optional mode where uploaded files are first submitted to a dedicated TotalAgility "Document Creator" process to obtain a Document ID.  The Chat Agent then receives the lightweight ID via the `DOCUMENT` input variable instead of the full base64 string, significantly reducing database load for large files.
- Added new environment variables: `PRELOAD_DOCUMENTS_AS_TOTALAGILITY_DOCS`, `TOTALAGILITY_DOCUMENT_CREATOR_PROCESS_ID`, `TOTALAGILITY_DOCUMENT_CREATOR_PROCESS_NAME`, `TOTALAGILITY_DOCUMENT_TYPE_ID`, `TOTALAGILITY_DOCUMENT_FILENAME_FIELD_ID`.

### Version 1.6
- **Notifications added to conversation history:** When a proactive notification is delivered via `POST /api/notifications`, the message is now also appended to the rolling conversation history (as "TotalAgility Agent Notification"). This gives the Chat Agent full context about what notifications the user has received when processing subsequent prompts.
- **Increased default conversation history size:** The default `CONVERSATION_HISTORY_MAX_ENTRIES` has been increased from `10` to `15` to provide the Chat Agent with more conversational context.
- **Debug command now shows conversation history:** The `debug` chat command now prints the current conversation history (with entry count) to both the Teams chat and the console log, in addition to the configuration summary.

### Version 1.7
- **Fixed document resubmission bug:** Previously, document-related input variables (`DOCUMENT`, `DOCUMENT_CONTENT`, `DOCUMENT_TYPE`, `DOCUMENT_FILENAME`) were sent on every request to the TotalAgility Agent — even when the user had not attached a file. This caused the same document to be reprocessed on every subsequent message turn.
- **Documents are now one-shot:** File attachments are only sent to the Chat Agent when the user explicitly uploads a file in the current turn. On text-only messages, the document variables are omitted entirely from the payload.
- **Refactored `callRestService()` signature:** The function now accepts an optional `documentInfo` object (or `null`) instead of separate `base64String`, `mimeType`, `fileName`, and `documentId` parameters. This makes it impossible to accidentally pass stale document data.
- **Removed noisy "No file attached" message:** The Chat Client no longer sends "No file attached. Processing your message…" on every text-only turn.

### Version 1.8
- **SSO session caching:** TotalAgility SSO session keys are now cached per user in `taAgent.js` and reused across requests, eliminating the overhead of requesting a new session for every API call.
- **Automatic 403 retry with loop prevention:** If a TotalAgility API call returns a 403 "Invalid Session ID" error (e.g. session timeout), the stale session is cleared, a fresh SSO login is performed, and the call is retried exactly **once**. If the retry also returns a 403, a user-friendly error message is displayed instead of retrying indefinitely — preventing infinite loops.
- **Graceful SSO failure handling:** If the SSO login itself fails (either on initial login or during a retry after session expiry), a user-friendly error message is returned to the Teams chat rather than an unhandled exception.
- **New auth-aware entry points:** Added `callRestServiceWithAuth()` and `createTotalAgilityDocumentWithAuth()` in `taAgent.js` which manage the full session lifecycle (obtain → use → detect expiry → refresh → retry). These are now the primary functions called by `teamsBot.js`.
- **Simplified `teamsBot.js`:** Removed the per-request `taSSOLogin()` call and the module-level `ssoKey` variable from `teamsBot.js`. The `handleMessageWithLoadingIndicator()` method no longer takes an `ssoKey` parameter — session management is fully encapsulated in `taAgent.js`.

---

## Microsoft 365 Agents Toolkit — Resources

The **Microsoft 365 Agents Toolkit** (formerly Teams Toolkit) is the primary toolchain for developing, debugging, and deploying Teams apps from Visual Studio Code. The following resources cover installation, project creation, and bot development:

| Resource | Description | Link |
|----------|-------------|------|
| **Install the Agents Toolkit** | Official guide for installing the Microsoft 365 Agents Toolkit extension in VS Code. | [learn.microsoft.com — Install Agents Toolkit](https://learn.microsoft.com/en-us/microsoftteams/platform/toolkit/install-agents-toolkit) |
| **Create a new project** | Step-by-step walkthrough for creating a new Teams app project using the Agents Toolkit. | [learn.microsoft.com — Create a new project](https://learn.microsoft.com/en-us/microsoftteams/platform/toolkit/create-new-project) |
| **Build a bot with Teams Toolkit** | Microsoft Learn training module covering how to create a Teams bot using the toolkit in VS Code. | [learn.microsoft.com — Teams Toolkit bot training](https://learn.microsoft.com/en-us/training/modules/teams-toolkit-vsc-create-bot/) |

## Example Resources

- Tutorial video: Creating a Basic AI Agent in TotalAgility — https://www.tungstendemocenter.com/items/creating-a-basic-ai-agent-in-totalagility

---

## Related Projects

| Project | Description |
|---------|-------------|
| [Agentic Design Patterns for TotalAgility](https://github.com/TungstenAutomationLabs/Agentic_Design_Patterns_For_TotalAgility) | Sample "Tool Use" Chat Agent built in TotalAgility, including examples of document conversion and SSO APIs. Use this companion repository to understand how to build the TotalAgility Agent processes that this Teams Chat Client connects to. |

