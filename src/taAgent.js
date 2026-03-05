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

async function callRestService(prompt_text, base64String, mimeType, sessionKey, fileName) {
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

    // Note: most values including temperature/seed/use_seed are now driven by
    // environment variables (see config.js).  Defaults are applied in code when
    // env settings are absent.
    // Note, removed previous syntax:    
    // ...(mimeType ? {
    // } : {}),
    // Since the Documents must be present when using create job and progress.
    //
    // 24/2/2026 - updated to use "/jobs/sync" only pass in the document using DOCUMENT_CONTENT, DOCUMENT_TYPE and DOCUMENT_FILENAME variables, and not use the Documents array. This is to allow the process to be called using "/jobs/sync" instead of "/jobs/progress".
    // This avoids a race condition where the sync job only saves the document contents to the database/storage after the job completes, which means the process can't launch associated jobs by passing in a document ID (if it does, it will get a document not found, as the document has not been saved yet). 
    // The chat process has been modified to use a separate step to create the TA document from the base64 content, and save it, passing back a document ID which can then be passed associated jobs.
    // Using "sync" means the process runs in memory, making it faster, and the audit logs etc. are only saved after the process completes. This requires dedicated saving of docs / data if subjobs need to lookup this data.
    let payload = {
        "ProcessId": "" + config.totalAgilityAgentId + "",
        "ProcessName": "" + config.totalAgilityAgentName + "",
        "JobInitialization": {
            "InputVariables": [
                {
                    "Id": "INPUT_PROMPT",
                    "Value": "" + prompt_text + ""
                },
                // temperature, use_seed and seed were originally hard‑coded.  They
                // can now be overridden via environment variables as defined in
                // config.js.  Default values are used otherwise.
                {
                    "Id": "TEMPERATURE",
                    "Value": (() => {
                        const t = parseFloat(config.totalAgilityTemperature);
                        return isNaN(t) ? 1 : t;
                    })()
                },
                {
                    "Id": "USE_SEED",
                    "Value": (() => {
                        const u = config.totalAgilityUseSeed;
                        if (typeof u === 'string') {
                            return u.toLowerCase() === 'true';
                        }
                        return u === undefined ? true : !!u;
                    })()
                },
                {
                    "Id": "SEED",
                    "Value": (() => {
                        const s = parseInt(config.totalAgilitySeed, 10);
                        return isNaN(s) ? 27535 : s;
                    })()
                },
                {
                    "Id": "DOCUMENT_CONTENT",
                    "Value": "" + base64String + ""
                },
                {
                    "Id": "DOCUMENT_TYPE",
                    "Value": "" + mimeType + ""
                },
                {
                    "Id": "DOCUMENT_FILENAME",
                    "Value": "" + fileName + ""
                }
            ]
        },
        "Documents": [],
        "VariablesToReturn": [
            {
                "VarId": "OUTPUT"
            }
        ],
        "StoreFolderAndDocuments": true,
        "ReturnOnlySpecifiedDocuments": true
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
        if (config.totalAgilityUseTestUser === "true") {
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