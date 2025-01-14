# Teams Chat UI for TotalAgility
This sample application illustrates connecting to TotalAgility Agents or Custom LLMs using REST APIs. 

Agents or Custom LLM processes in TotalAgility are synchronous process built with the TotalAgility low code designer.

They provide a standard interface for both chat, document and API interactions. 

This Teams application front end calls the Agent in TotalAgility and displays the response. All logic for automation, data, knowledge base and LLM calls are managed in the TotalAgility Agent, opening access from teams to a wide range of automation capabilities including: 
- Workflows & business process management 
- Case management
- SLA monitoring & management
- Worktype modelling, status tracking, process events and milestones
- Agent governace & auditing
- IDP
- RPA
- Document Generation
- eSigning
- Document generation & transformation 
- Knowledge base creation & management
- Automated knowledge discovery 

By using Agents in TotalAgility to lauch cases, workflows and automation flows, the Agent is the gateway to a wider world of automation.

This code sample acts as a proxy to the TotalAgility Agent interface, passing in calls and conversation history, and displaying the responses.

As it is implemented using the Microsoft Bot framework, they application can be used in other chat and voice environments, including:
- Webchat
- Facebook Messenger
- Alexa
- WhatsApp
- etc.

The logic in the Teams code is minimal, providing:
- Connectivity to the TotalAgility API
- Maintaining conversation history (which is passed to each TotalAgility API call, thus maintaining state over multiple invocations of the API)
- Display of user feedback in the form of loading messages and "typing" feedback

The code is implemented in node.js and can be run locally or deployed to a Mircosoft 365 environment. 

The base package is based from the VS Code template:
Teams > Development > Create New App > Bot

Note that the files in the env folder need to be updated to match your environment. Specifically note the addition of 2 new fields:
TOTALAGILITY_ENDPOINT=
TOTALAGILITY_API_KEY=
