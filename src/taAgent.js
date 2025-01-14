/*************************************************************************/
/*      FUNCTIONS FOR CALLING TOTALAGILITY PROCESSES AND WORKFLOWS       */
/*************************************************************************/

const config = require("./config");

function tester(text) {
    console.log("Calling TotalAgility on " + config.totalAgilityEndpoint);
    console.log("Calling TotalAgility on " + config.totalAgilityApiKey);
    return "Hello " + text;
}

async function callRestService(prompt_text) {
    console.log("callRestService() called with: " + prompt_text);
    
    let return_response = "xxxx";
    console.log("Calling TotalAgility on " + config.totalAgilityEndpoint);
    const url = config.totalAgilityEndpoint;

    // Note some hard coded values for the process ID, seed etc.
    let payload = {
        "ProcessId": "04B1CCBD831589A837633D4DBB013EC3",
        "ProcessName": "Intent Router Agent",
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
    const headers = {
        'Content-Type': 'application/json',
        'Authorization': '' + config.totalAgilityApiKey + '' // Replace with your token
    };


    try {
        let response = await fetch(url, {
            method: 'POST',
            headers: headers,
            body: JSON.stringify(payload)
        });

        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
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
        return_response = 'Error: ' + error;
    }
    return return_response;
}

// Export the functions
module.exports = {
    tester,
    callRestService
};