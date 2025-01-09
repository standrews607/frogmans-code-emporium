import * as msal from '@azure/msal-node';
import axios from 'axios';
import dotenv from 'dotenv';
dotenv.config();

/*
* Note, use NPM to install the following packages:
* npm install @azure/msal-node
* npm install axios
* npm install dotenv
* Ref documentation: https://learn.microsoft.com/en-us/entra/identity-platform/quickstart-console-app-nodejs-acquire-token
* Ref documentation: https://learn.microsoft.com/en-us/graph/api/user-list?view=graph-rest-1.0&tabs=http
* Ref documentation: https://learn.microsoft.com/en-us/entra/identity-platform/tutorial-v2-nodejs-console
* I have of course modifed this slightly from the example having compressed everything into a single file.
*/

const apiConfig = {
    uri: process.env.GRAPH_ENDPOINT + '/v1.0/users',
};

// Auth variables
const Tenant_Id = process.env.Tenant_Id;
const Client_Id = process.env.Client_Id;
const Client_Secret = process.env.Client_Secret;


async function getToken() {

    const msalConfig = {
        auth: {
            clientId: Client_Id,
            authority: `https://login.microsoftonline.com/${Tenant_Id}`,
            clientSecret: Client_Secret,
        }
    };
    const cca = new msal.ConfidentialClientApplication(msalConfig);
    const tokenRequest = {
        scopes: [ 'https://graph.microsoft.com/.default' ],
    };
    const tokenResponse = await cca.acquireTokenByClientCredential(tokenRequest);

    // Return token
    return tokenResponse.accessToken;

}

async function callApi(endpoint, accessToken) {

    const options = {
        headers: {
            Authorization: `Bearer ${accessToken}`
        }
    };

    console.log('request made to web API at: ' + new Date().toString());

    try {
        const response = await axios.get(endpoint, options);
        return response.data;
    } catch (error) {
        console.log(error)
        return error;
    }
};

async function main() {
    const authToken = await getToken();
    const users = await callApi(apiConfig.uri, authToken);
    users.value.forEach(user => {
        console.log(user.userPrincipalName);
    });
}

main();