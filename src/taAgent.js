/*************************************************************************/
/*      FUNCTIONS FOR CALLING TOTALAGILITY PROCESSES AND WORKFLOWS       */
/*************************************************************************/

const config = require("./config");
const { TeamsInfo } = require('botbuilder');

function tester(text) {
    console.log("Calling TotalAgility on " + config.totalAgilityEndpoint);
    console.log("Calling TotalAgility on " + config.totalAgilityApiKey);
    return "Hello " + text;
}

async function callRestService(prompt_text, base64String, mimeType, sessionKey) {
    console.log("callRestService() called with: " + prompt_text);

    if (base64String) {
        console.log("File attached with size: " + base64String.length);
    } else {
        base64String = "";
    }

    if (mimeType) {
        console.log("File MIME type: " + mimeType);
    } else {
        mimeType = "";
    }

    let return_response = "xxxx";
    console.log("Calling TotalAgility on " + config.totalAgilityEndpoint + "/jobs/sync");
    const url = config.totalAgilityEndpoint + "/jobs/sync";

    // Note some hard coded values for the process ID, seed etc.
    let payload = {
        "ProcessId": "" + config.totalAgilityAgentId + "",
        "ProcessName": "" + config.totalAgilityAgentName + "",
        "JobInitialization": {
            "InputVariables": [
                {
                    "Id": "INPUT_PROMPT",
                    "Value": "" + prompt_text + ""
                },
                {
                    "Id": "TEMPERATURE",
                    "Value": 0.8
                },
                {
                    "Id": "USE_SEED",
                    "Value": true
                },
                {
                    "Id": "SEED",
                    "Value": 27535
                }
            ]
        },
        ...(mimeType ? {
            "Documents": [
                {
                    "MimeType": "" + mimeType + "",
                    "RuntimeFields": [],
                    "FolderId": "",
                    "DocumentTypeId": "",
                    "FolderTypeId": "",
                    "Base64Data": "" + base64String + "",
                    "DocumentTypeName": "",
                    "DocumentGroupId": "",
                    "DocumentGroupName": ""
                }
            ]
        } : {}),
        "VariablesToReturn": [
            {
                "VarId": "OUTPUT"
            }
        ]
    };

    /*
    const headers = {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer YOUR_ACCESS_TOKEN' // Replace with your token
    };
    */
    /* Without SSO:
     const headers = {
         'Content-Type': 'application/json',
         'Authorization': '' + config.totalAgilityApiKey + '' // Replace with your token
     };
     */
    const headers = {
        'Content-Type': 'application/json',
        'Authorization': '' + sessionKey + ''
    };
    
    try {
        let response = await fetch(url, {
            method: 'POST',
            headers: headers,
            body: JSON.stringify(payload)
        });

        if (!response.ok) {
            // Debug:
            // throw new Error(`HTTP error! status: ${response.status} \n\n URL: ${url} \n\n Payload: ${JSON.stringify(payload)} \n\n Headers: ${JSON.stringify(headers)}`);
            throw new Error(`HTTP error! status: ${response.status} \n\n URL: ${url} \n\n`);
        }

        let data = await response.json();
        console.log('JobId:', data.JobId); // Log the JobId from the response

        // Check if ReturnedVariables exist and log their values
        if (data.ReturnedVariables && data.ReturnedVariables.length > 0) {
            data.ReturnedVariables.forEach(variable => {
                console.log('Returned Variable Value:', variable.Value);
                return_response = variable.Value;
                //response = "Hello World";
            });
            //response = data.ReturnedVariables[0].Value;

        } else {
            console.log('No Returned Variables found.');
            return_response = 'No Returned Variables found.';
        }

    } catch (error) {
        console.error('Error: ', error);
        return_response = 'Error: ' + error + "\nPayload: " + JSON.stringify(payload);
        // return_response = 'Error: ' + error + ' using url: ' + url + ' and key ' + config.totalAgilityApiKey + "\nPayload: " + JSON.stringify(payload);
    }
    return return_response;
}

async function taSSOLogin(context) {
    try {
        let sessionID = "";
        const ssoUrl = config.totalAgilityEndpoint + "/users/sessions/single-sign-on";

        let ssoPayload = {
            "UserId": ""
        }
        if (config.totalAgilityUseTestUser === "true" ) {
            // Use test user SSO login:
            ssoPayload.UserId = config.totalAgilityTestUserName;
        } else {
            // Production SSO login:
            const userInfo = await getCurrentUserIdAndEmail(context);
            ssoPayload.UserId = userInfo.email; // Assumes the user's email is their TA UserID
        }

        // Use the API key as the authorization for SSO login:
        const ssoHeaders = {
            'Content-Type': 'application/json',
            'Authorization': '' + config.totalAgilityApiKey + ''
        };

        // Get the SSO token
        let ssoResponse = await fetch(ssoUrl, {
            method: 'POST',
            headers: ssoHeaders,
            body: JSON.stringify(ssoPayload)
        });

        if (!ssoResponse.ok) {
            throw new Error(`HTTP error! status: ${ssoResponse.status} \n\n URL: ${ssoUrl} \n\n Use Test user: ${config.totalAgilityUseTestUser} \n\n Test UserID: ${config.totalAgilityTestUserName} \n\n Payload: ${JSON.stringify(ssoPayload)}  `);
        }

        let ssoData = await ssoResponse.json();
        sessionID = ssoData.SessionId;
        return sessionID;

    } catch (error) {
        console.error('SSO Login Error: ', error);
        // return 'Error during SSO Login: ' + error;
        throw error; // Bubble up the error
    }
}

/**
 * Gets the current user's Teams ID and email address.
 * @param {TurnContext} context - The turn context from the bot.
 * @returns {Promise<{id: string, email: string}>}
 */
async function getCurrentUserIdAndEmail(context) {
  try {
    // Fetch the member info using the user's Teams ID
    const member = await TeamsInfo.getMember(context, context.activity.from.id);
    // member.id is the Teams user ID, member.email is the user's email
    return {
      id: member.id,
      email: member.email
    };
  } catch (error) {
    // Handle the case where the member info can't be retrieved
    console.error('Failed to get user info:', error);
    return null;
  }
}

// Export the functions
module.exports = {
    tester,
    callRestService,
    taSSOLogin
};