# Teams Chat UI for TotalAgility Agents
This sample application illustrates connecting to TotalAgility Agents or Custom LLMs using REST APIs. 

Provides a lightweight MS Teams / Bot Framework UI to front a TotalAgility Automation Agent, proxying user interactions via Teams to the Agent running in TotalAgility, then displaying the Agent's responses.

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

This code sample acts as a proxy to the TotalAgility Agent interface, passing in calls and conversation history, and displaying the responses. As such it only implements basic Teams functionality leaving the core chat / agent logics to the processes running in TotalAgility. 

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
```TOTALAGILITY_ENDPOINT=```
```TOTALAGILITY_API_KEY=```

The TotalAgility endpoint is the TA Rest / OpenAPI create jobs sync endpoint. For more details see the TotalAgility documentation. 
```https://{{your_tenant}}.dev.kofaxcloud.com/services/sdk/v1/jobs/sync```

The main API call is managed in the ```src/taAgent.js``` file. Update this file with the corresponding details for the TA Agent you want to call. Note that this API call uses a hard coded "seed" for consistent responses, but this can be changed or omitted as desired. 

The pattern I've used is to have a single controlling Agent process, in my environment call the "Intent Router Agent".
This process, implemented as a "Custom LLM" or "Agent Process" in TotalAgility (the name used will depend on the version of TotalAgility you are using), in turn calls other agent processes depending on the prompt sent. 

The "Intent Router Agent" makes use of an LLM step to evalute the incoming prompt, and map this to one of the available actions or sub-agents. 
The incoming prompt is combined with a prompt guiding the LLM on how to chose the intent or next action. This system prompt include a registry or list of available next actions / agents, including their ProcessIDs in the data structure. The intent evaluator step returns json data mapped to a TotalAgility Data Model embedding the details of the sub process / agent to call, and the prompt to pass in. 

As all agents have the same interface, the registry of available actions can be 100s of different agents, each provided with a description allowing the managing agent / intent router agent to determine an appropriate next step. 

The final step is an evaluation step, which determines if the agent is ready to respond, of if the agent should repeat the process to find additonal information, data or take another action (this allows the agent to undertake multiple steps to complete a task or goal, for example searcing both the internet and a knowledge base for content to use in a RAG pattern). 

![Example intent router process](intent-router-agent-process.png)

