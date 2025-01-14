const { TeamsActivityHandler, TurnContext } = require("botbuilder");

const TotalAgilityAgent = require('./taAgent.js');

const Utils = require('./utils.js');

let messageArray = []; // An array to hold the current chat history. 
const messageArrayMaxSize = 10; // Max number of messages to retain in the chat history.

class TeamsBot extends TeamsActivityHandler {
  constructor() {
    super();

    this.onMessage(async (context, next) => {
      context.sendActivity(Utils.getRandomLoadingMessage()); // Should be displayed without waiting for the handleMessageWithLoadingIndicator to finish first
      await context.sendActivities([{ type: 'typing' }]); // Display the "typing" animation. Including twice as this seem to ensure it is consistenly diplayed
      await this.handleMessageWithLoadingIndicator(context); // Call the TA Agent / API
      await next();
    });


    // Listen to MembersAdded event, view https://docs.microsoft.com/en-us/microsoftteams/platform/resources/bot-v3/bots-notifications for more events
    this.onMembersAdded(async (context, next) => {
      const membersAdded = context.activity.membersAdded;
      for (let cnt = 0; cnt < membersAdded.length; cnt++) {
        if (membersAdded[cnt].id) {
          /* Uncomment to show an initial "welcome" message:
          await context.sendActivity(
            `Hi there! I'm TotalAgility, your Intelligent Automation Agent.`
          );
          */
          break;
        }
      }
      await next();
    });
  }

  async handleMessageWithLoadingIndicator(context) {
    await context.sendActivities([{ type: 'typing' }]); // Display the "typing" animation. Including twice as this seem to ensure it is consistenly diplayed
    console.log("Running with Message Activity.");
    const removedMentionText = TurnContext.removeRecipientMention(context.activity);
    const userRequest = removedMentionText.toLowerCase().replace(/\n|\r/g, "").trim();
    saveMsg("User", userRequest); // Save a copy of the received message 
    
    try {
      // Send just the most recent prompt to TotalAgility: 
      // let agentResponse = await TotalAgilityAgent.callRestService(userRequest);

      // Send the whole conversation history to  the TA Agent / API.
      let agentResponse = await TotalAgilityAgent.callRestService(Utils.renderConversationHistoryMarkdown(messageArray));
      await context.sendActivity(agentResponse); // Send the response to the user
      saveMsg("TotalAgility Bot", agentResponse); // Save a copy of the reply message

    } catch (error) {
      await context.sendActivity(`An error occurred: ${error.message}`);
    }
    // Debug the contents of the message array by sending to the user:
    // await context.sendActivity(`Current content of messageArray: \n\n` + Utils.renderConversationHistoryMarkdown(messageArray));
  }
}


function saveMsg(actor, message) {
  // Add the new element to the array
  // messageArray.push(message);

  messageArray.push({ speaker: actor, message: message });

  // Check if the array exceeds the maximum size
  if (messageArray.length > messageArrayMaxSize) {
    // Remove the oldest entry (first element)
    messageArray.shift();
  }
}


module.exports.TeamsBot = TeamsBot;
