// move to package.json later
const { TeamsInfo } = require('botbuilder');
const { AdaptiveCardHelper } = require('./adaptiveCardHelper.js');
const { CardResponseHelpers } = require('./cardResponseHelpers.js');
const { EmailHelper } = require('./emailHelper.js');
const { SimpleGraphClient } = require('./simpleGraphClient.js');
const axios = require('axios');

class AuthHelper {

    /**
     * @param {string} tenant AAD Tenant ID of current user
     * Return a token containing app-only permissions for admin-approved scopes: 
     * EduRoster.Read.All, Mail.Read, Mail.Send, User.Read.All (as of 7/23/20) 
     * */
    static async getUserToken(context) {
        const tenant = context.activity.conversation.tenantId;
        const grantType = 'client_credentials';
        const scopeURL = 'https://graph.microsoft.com/.default';
        // temp hard code for AAD app details 
        const AADappID = 'ae7e6f3c-c15e-4d80-bba0-5ba655e00e4d';  
        const clientSecret = '_~3hiB4dzE14Gs~9ll7F.bO7u3sGWqan~F';

        const params = new URLSearchParams();
        params.append('grant_type', grantType);
        params.append('client_id', AADappID);
        params.append('scope', scopeURL);
        params.append('client_secret', clientSecret);

        const tokenResponse = await axios({
            method: 'post',
            url: `https://login.microsoftonline.com/${tenant}/oauth2/v2.0/token`,
            headers: {
                'Content-Type': 'application/x-www-form-urlencoded'
            },
            data: params
            
        })
        .catch(function (error) {
            if (error.response) {
            // The request was made and the server responded with a status code
            // that falls out of the range of 2xx
            console.log(error.response.data);
            console.log(error.response.status);
            console.log(error.response.headers);
            } else if (error.request) {
            // The request was made but no response was received
            // `error.request` is an instance of XMLHttpRequest in the browser and an instance of
            // http.ClientRequest in node.js
            console.log(error.request);
            } else {
            // Something happened in setting up the request that triggered an Error
            console.log('Error', error.message);
            }
            console.log('CONFIG: ')
            console.dir(error.config);
        });
        return tokenResponse.data.access_token;
    }

    /**
     * @param {Object} context Current chat context 
     * Handle the 'Email' command with user authentication (ie context-based token retrieval)
     * */
    static async handleEmailAuthCommand(context) {
        const token = await this.getUserToken(context);
        console.log(token);
        // Retrieve user email after authentication
        const senderEmail = await EmailHelper.listEmailAddress(context, token);

        // Retrieve team name 
        var teamName = '';
        var classID = '';
        var teamDetails;
        try {
            teamDetails = await TeamsInfo.getTeamDetails(context);
        } catch (e) {
            console.log(e);
            throw e;
        }
        if (teamDetails) {
            teamName = `${ teamDetails.name}`;
            classID = teamDetails.aadGroupId;
        }

        // Form recipient group ID
        const recipientGroupID = teamName+' Parents & Guardians';

        // Pull in EDU data from Microsoft Graph.
        const client = new SimpleGraphClient(token);

        var classIDsAndRoles = '';
        var eduRoster;
        try {
            eduRoster = await client.getEduRoster(classID); // Return id and primaryRole if it exists
        } catch (e) {
            console.log(e);
            throw e;
        }
        if (eduRoster) {
            classIDsAndRoles = eduRoster.value;
        }

        // Store student IDs
        const studentIDs = [];
        for (let memberIDAndRole of classIDsAndRoles) {
            if (memberIDAndRole.primaryRole) {
                if (memberIDAndRole.primaryRole === 'student') {
                    studentIDs.push(memberIDAndRole.id);
                }
            } 
        }

        var formattedContactNames=[];
        var emailListString=''

        for (let studentID of studentIDs) {
            // Retrieve parent/guardian emails and names as a subset of a student's relatedContacts
            // graph query does not allow exclusive select for relatedContacts
            const relatedContactsAndId = await client.getRelatedContactsAndId(studentID);
            if (relatedContactsAndId){
                if (relatedContactsAndId.relatedContacts.length > 0) {
                    const relatedContacts = relatedContactsAndId.relatedContacts; // student's related contacts
                    for (let i=0; i < relatedContacts.length; i++) {
                        if (relatedContacts[i].relationship === 'parent' || relatedContacts[i].relationship === 'guardian') {
                            // contactEmails.push(relatedContacts[i].emailAddress);
                            emailListString+=relatedContacts[i].emailAddress + ",";
                            // correct JSON format for Adaptive Card Choice Set 
                            formattedContactNames.push({
                                title: relatedContacts[i].displayName,
                                value: relatedContacts[i].emailAddress
                            });

                        }
                    }
                }
            } 
        }

        console.log(formattedContactNames);
        console.log(emailListString);

        const adaptiveCard = AdaptiveCardHelper.createAdaptiveCardEditor(senderEmail, recipientGroupID, formattedContactNames, emailListString);
        return CardResponseHelpers.toTaskModuleResponse(adaptiveCard);
    }
}
exports.AuthHelper = AuthHelper;