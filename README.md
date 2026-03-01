
# Teams Chat UI for TotalAgility Agents

This repository is a lightweight Microsoft Teams front-end (using the Microsoft Bot Framework) that proxies chat interactions to TotalAgility Agents or custom LLM-backed Agents via REST APIs. It displays agent responses in Teams while leaving the automation, data access, and LLM logic inside TotalAgility.

![Example screenshot of an interaction with a TotalAgility AI Agent](img/TotalAgility-Teams-App-Screenshot.png)

Key capabilities:
- Proxy Teams chat to a TotalAgility Agent (or Custom LLM) using REST/OpenAPI
- Preserve and pass conversation history to the Agent to maintain context
- Support for sending documents (file upload) to the Agent (see Version notes)
- Minimal Teams-side logic so the Agent process in TotalAgility can own orchestration, knowledge access and calling other sub-agents

Why use this sample
- It demonstrates how Teams can be used as a UI for complex automation and LLM-driven processes implemented in TotalAgility (workflows, case management, RPA, IDP, knowledge base lookups, document generation, eSigning, etc.).

Architecture summary
- The Teams app acts as a thin proxy: it accepts user messages, forwards them (with conversation history) to a TotalAgility Agent endpoint, then renders the Agent's response in Teams.
- The recommended pattern is an "Intent Router Agent" in TotalAgility that receives the prompt, evaluates intent via an LLM step, and routes to specific sub-agents/processes (each with the same Agent interface).

Supported channels
- Microsoft Teams (primary)
- Any Bot Framework channel (Webchat, Facebook Messenger, WhatsApp, Alexa, etc.) with minimal changes.

Project structure and where to make edits
- appPackage/: Teams app manifest and package. Example prompts and app configuration live in `appPackage/manifest.json`.
- src/: core application code
	- `src/index.js` — app entry point and server bootstrap
	- `src/config.js` — configuration loader and environment handling
	- `src/taAgent.js` — main TotalAgility API integration and the primary place to change how the app calls your Agent (seed usage, request shape, headers, error handling)
	- `src/teamsBot.js` — Teams/Microsoft Bot framework adapter and conversation turn handling
	- `src/taAgent.js` — primary API wrapper for calling TotalAgility (modify this to change request/response handling)
	- `src/utils.js` — helper functions including loading/typing feedback messages (customize UI text here)
- env/: environment template files — update these to match your tenant, keys and Agent details
- infra/: infrastructure deployment scripts (Azure Bicep templates used for provisioning, if desired)
- devTools/: developer utilities (Teams App Tester, etc.)

Important files to edit for common tasks
- Change the Agent integration or modify request details: `src/taAgent.js`
- Adjust loading/typing UI and helper utilities: `src/utils.js`
- For Teams-specific behaviour or message formatting, update `src/teamsBot.js` and `src/index.js`.
- Keep environment-specific secrets out of source control. Use the `env/` templates and local environment variables when running locally.

Environment variables / configuration
Place your environment values in the files under `env/` (or in your system environment). The sample adds the following keys:

```
TOTALAGILITY_ENDPOINT=
TOTALAGILITY_API_KEY=
TOTALAGILITY_AGENT_NAME=
TOTALAGILITY_AGENT_ID=
TOTALAGILITY_TEST_USERNAME=
TOTALAGILITY_USE_TEST_USER=
```

Notes on these values
- `TOTALAGILITY_ENDPOINT` — base TotalAgility REST/OpenAPI endpoint, e.g. `https://{{your_tenant}}.dev.kofaxcloud.com/services/sdk/v1`
- `TOTALAGILITY_API_KEY` — API key used to authenticate calls to TotalAgility
- `TOTALAGILITY_AGENT_NAME` — the process name of the Agent in TotalAgility
- `TOTALAGILITY_AGENT_ID` — the process ID (often visible in the TotalAgility Designer edit URL)
- `TOTALAGILITY_TEST_USERNAME` & `TOTALAGILITY_USE_TEST_USER` — override SSO behaviour to force a test TA user (useful for development)

Behavior notes preserved from the original sample
- The main API call is managed in `src/taAgent.js`. The sample uses a hard-coded "seed" for consistent responses; remove or change this if you want nondeterministic LLM outputs.
- Loading messages can be configured in `src/utils.js`.
- Example prompts are available in `appPackage/manifest.json`.

How the Intent Router pattern works (summary from original README)
- A controlling "Intent Router Agent" evaluates incoming prompts using an LLM step and maps them to available actions / sub-agents. The router provides a registry of available agents (with ProcessIDs) and returns JSON mapped to the TotalAgility data model describing which sub-process to invoke and which prompt to send.
- Because all agents expose a common interface, the registry can include many agents; the router chooses the best match and may iterate multiple steps (search KB, call external APIs, gather documents) before returning a final response.

Example resources
- Tutorial video: Creating a Basic AI Agent in TotalAgility — https://www.tungstendemocenter.com/items/creating-a-basic-ai-agent-in-totalagility

Version notes 
### Version 1.1
- Added the ability to upload files and send these to the TotalAgility Agent for processing. This sample uses TotalAgility 25.2 where the Agent interface accepts TotalAgility Documents (sent as base64 strings) to the Jobs sync API.

### Version 1.2
- Added settings to SSO a user into TotalAgility based on their email address from their Teams login.

*Note:* the code assumes the user's email address (from their MS Teams login) is their user ID in TotalAgility. To override this, specify a test user in the environment and set the `TOTALAGILITY_USE_TEST_USER` flag to `true`.

Environment variables for SSO testing:

```
TOTALAGILITY_TEST_USERNAME=my_ta_test_account@test.com
TOTALAGILITY_USE_TEST_USER=true
```

### Running and developing locally
- Install dependencies:

```bash
npm install
```

- Run the app locally (Teams development): use the Visual Studio Code tasks in this workspace or run the equivalent npm scripts. Example tasks present in the workspace:
	- `Start Teams App (Test Tool)` — starts the app using the Test Tool configuration
	- `Start Teams App Locally` — runs local tunnel, provision, deploy and starts the app

- Common npm scripts (available in `package.json`):

```bash
npm run dev:teamsfx        # start locally for Teams development
npm run dev:teamsfx:testtool  # start using the Test Tool flow
```

### Editing tips
- To change how the app calls TotalAgility (payload, headers, error handling), update `src/taAgent.js`.
- To change user-visible loading/typing messages, update `src/utils.js`.
- For Teams-specific behaviour or message formatting, update `src/teamsBot.js` and `src/index.js`.
- Keep environment-specific secrets out of source control. Use the `env/` templates and local environment variables when running locally.

### Deployment and infra
- The `infra/` folder contains Azure Bicep templates to help provision cloud resources if you want to deploy to Azure.



