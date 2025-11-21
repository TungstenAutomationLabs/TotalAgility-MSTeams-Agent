const { TeamsActivityHandler, TurnContext } = require("botbuilder");

const TotalAgilityAgent = require('./taAgent.js');
//const config = require("./config"); // Uncomment for some debug steps
const Utils = require('./utils.js');

// Utils for file handling:
const axios = require('axios');
const fs = require('fs');
const path = require('path');

let messageArray = []; // An array to hold the current chat history. 
const messageArrayMaxSize = 10; // Max number of messages to retain in the chat history.
let ssoKey = ""; // Variable to hold the SSO key

class TeamsBot extends TeamsActivityHandler {
  constructor(userState, ssoKeyAccessor) {
  // constructor() {
    super();
    this.userState = userState;
    this.ssoKeyAccessor = ssoKeyAccessor;

    this.onMessage(async (context, next) => {

      await context.sendActivity(Utils.getRandomLoadingMessage()); // Should be displayed without waiting for the handleMessageWithLoadingIndicator to finish first
      //await context.sendActivities([{ type: 'typing' }]);

      // Try to get the SSO key from state
      ssoKey = await this.ssoKeyAccessor.get(context);

      if (!ssoKey) {
        await context.sendActivity(`Signing into TotalAgility...`);
        await context.sendActivities([{ type: 'typing' }]);
        // If not present, call your async function to get it
        ssoKey = await TotalAgilityAgent.taSSOLogin(context); // Get the SSO Key from TotalAgility
    
        // Store the SSO key in state for future use
        await this.ssoKeyAccessor.set(context, ssoKey);
        await context.sendActivity(`TotalAgility sign-in complete.`);
        // Debug:
        //await context.sendActivity(`TotalAgility sign-in complete... SSO Key: ${ssoKey}`);
      }

      await context.sendActivities([{ type: 'typing' }]); // Display the "typing" animation. Including twice as this seem to ensure it is consistenly diplayed
      await this.handleMessageWithLoadingIndicator(context,ssoKey); // Call the TA Agent / API
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

  async handleMessageWithLoadingIndicator(context,ssoKey) {
    await context.sendActivities([{ type: 'typing' }]); // Display the "typing" animation. Including twice as this seem to ensure it is consistenly diplayed
    console.log("Running with Message Activity.");


    try {

      const removedMentionText = TurnContext.removeRecipientMention(context.activity);
      const userRequest = removedMentionText.toLowerCase().replace(/\n|\r/g, "").trim();
      saveMsg("User", userRequest); // Save a copy of the received message

      // Init var to hold base64 string if a file is attached:
      let base64String = "";
      let mimeType = "";

      // Check for presence of a file: 
      if (context.activity.attachments && context.activity.attachments.length > 0) {

        // File handling code here - process each attachment:
        for (let i = 0; i < context.activity.attachments.length; i++) {
          const attachment = context.activity.attachments[i];

          if (attachment.contentType === 'application/vnd.microsoft.teams.file.download.info') { // This indicates a file attachment

            const downloadUrl = attachment.content.downloadUrl;
            const contentUrl = attachment.contentUrl;
            const contentType = attachment.contentType;
            const fileName = attachment.name;

            mimeType = getMimeType(attachment.content.fileType); // Get MIME type based on file extension

            // Debug - inform user of received file:
            //const fileMessage = `Received file: ${fileName} (Type: ${contentType}, URL: ${contentUrl})`;
            //await context.sendActivity(fileMessage);
            //saveMsg("TotalAgility Bot", fileMessage); // Save a copy of the reply message

            // Update the user:
            await context.sendActivity(`Uploading file ${fileName}.`);
            await context.sendActivities([{ type: 'typing' }]);

            // Download the file as a buffer
            const response = await axios.get(downloadUrl, { responseType: 'arraybuffer' });
            const fileBuffer = Buffer.from(response.data);

            // Convert file buffer contents to a base64 string, to send to TotalAgility:
            base64String = fileBuffer.toString('base64'); 

            // Debug:
            //await context.sendActivity(`File ${fileName} received and converted to base64! Sending to Agent: ${config.totalAgilityAgentName} \n ${base64String.substring(0, 100)}...`); // Send first 100 characters only for brevity
            await context.sendActivity(`File ${fileName} received.`);
            await context.sendActivities([{ type: 'typing' }]); // Show typing indicator while processing

          } else {
            // Skip non-file attachments and messages:
            // await context.sendActivity(`Attachment ${i + 1} is not a file download info type.`);
          }
        }

        // Debug:
        // await context.sendActivity(JSON.stringify(context)); // Debug - show the first attachment object
      } else {
        // No file attached, proceed with normal message processing
        await context.sendActivity(`No file attached. Processing your message...`);
      }
      // End file handling code.

      // Send just the most recent prompt to TotalAgility: 
      // let agentResponse = await TotalAgilityAgent.callRestService(userRequest);

      // Send the whole conversation history to  the TA Agent / API.
      let agentResponse = await TotalAgilityAgent.callRestService(Utils.renderConversationHistoryMarkdown(messageArray), base64String, mimeType, ssoKey); // Pass the base64 string if a file was attached
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
  messageArray.push({ speaker: actor, message: message });

  // Check if the array exceeds the maximum size
  if (messageArray.length > messageArrayMaxSize) {
    // Remove the oldest entry (first element)
    messageArray.shift();
  }
}


function getMimeType(ext) {
    // Remove leading dot if present and convert to lowercase
    ext = ext.replace(/^\./, '').toLowerCase();

    // TODO - align to list of supported types in TotalAgility & reject if unsupported (currently handled TA side)

    // Common extension to MIME type mapping
    const mimeTypes = {
        'jpg': 'image/jpeg',
        'jpeg': 'image/jpeg',
        'png': 'image/png',
        'gif': 'image/gif',
        'bmp': 'image/bmp',
        'webp': 'image/webp',
        'svg': 'image/svg+xml',
        'ico': 'image/x-icon',
        'tiff': 'image/tiff',
        'pdf': 'application/pdf',
        'txt': 'text/plain',
        'html': 'text/html',
        'htm': 'text/html',
        'css': 'text/css',
        'tif': 'image/tiff',
        'tiff': 'image/tiff',
        'js': 'application/javascript',
        'json': 'application/json',
        'xml': 'application/xml',
        'csv': 'text/csv',
        'zip': 'application/zip',
        'rar': 'application/vnd.rar',
        'tar': 'application/x-tar',
        'gz': 'application/gzip',
        'mp3': 'audio/mpeg',
        'wav': 'audio/wav',
        'ogg': 'audio/ogg',
        'mp4': 'video/mp4',
        'avi': 'video/x-msvideo',
        'mov': 'video/quicktime',
        'wmv': 'video/x-ms-wmv',
        'flv': 'video/x-flv',
        'mkv': 'video/x-matroska',
        'doc': 'application/msword',
        'docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        'xls': 'application/vnd.ms-excel',
        'xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'ppt': 'application/vnd.ms-powerpoint',
        'pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        // Add more as needed
    };

    return mimeTypes[ext] || null;
}


module.exports.TeamsBot = TeamsBot;
