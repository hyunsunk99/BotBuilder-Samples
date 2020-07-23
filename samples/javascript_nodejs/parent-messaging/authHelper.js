// move to package.json later
const { TeamsInfo } = require('botbuilder');
const { AdaptiveCardHelper } = require('./adaptiveCardHelper.js');
const { CardResponseHelpers } = require('./cardResponseHelpers.js');
const { EmailHelper } = require('./emailHelper.js');
const { SimpleGraphClient } = require('./simpleGraphClient.js');
const axios = require('axios');

class AuthHelper {
    static async handleEmailAuthCommand(context, action, connectionName) {

        // temp hard code
        const tenant = context.activity.conversation.tenantId;
        const grantType = 'client_credentials';
        const AADappID = 'ae7e6f3c-c15e-4d80-bba0-5ba655e00e4d';  
        const scopeURL = 'https://graph.microsoft.com/.default';
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

        const token = tokenResponse.data.access_token;
        console.log(token);
        // SIMPLE EMAIL WORKING  
        await EmailHelper.sendMail(context, token, 'heidik87@gmail.com');

        // Retrieve user email after authentication
        const senderEmail = await EmailHelper.listEmailAddress(context, token);

        const client = new SimpleGraphClient(token);	

        // Retrieve team name 
        var teamName = '';
        var teamDetails;
        try {
            teamDetails = await TeamsInfo.getTeamDetails(context);
        } catch (e) {
            console.log(e);
            throw e;
        }
        if (teamDetails) {
            teamName = `${ teamDetails.name}`;
        }

        // Return either a sign-in modal (user not verified) or the email template (verified)

        // Retrieve recipient group ID (hard code for now)
        const recipientGroupID = teamName+' Parents & Guardians';
        const teamID = teamDetails.aadGroupId;

        // Retrieve team members 
        var memberIds = [];
        var chatMembers;
        try {
            chatMembers = await TeamsInfo.getPagedMembers(context);
        } catch (e) {
            console.log(e);
            throw e;
        }
        if (chatMembers) {
            for (let i = 0; i <teamDetails.memberCount;i++) {
                memberIds.push(chatMembers.members[i].aadObjectId);
            }
        }

        var parentEmails = [];
        for (let memberId of memberIds) {
            // graph query does not allow exclusive select for relatedContacts
            const relatedContactsAndId = await client.getRelatedContactsAndId(memberId);
            if (relatedContactsAndId){
                if (relatedContactsAndId.relatedContacts.length > 0) {
                    const relatedContacts = relatedContactsAndId.relatedContacts; // student's related contacts
                    for (let i=0; i < relatedContacts.length; i++) {
                        if (relatedContacts[i].relationship === 'parent' || relatedContacts[i].relationship === 'guardian') {
                            parentEmails.push(relatedContacts[i].emailAddress);
                        }
                    }
                }
            } 
        }
        console.log(parentEmails);

        const adaptiveCard = AdaptiveCardHelper.createAdaptiveCardEditor(senderEmail, recipientGroupID);
        return CardResponseHelpers.toTaskModuleResponse(adaptiveCard);
    }

    static async handleSignOutCommand(context, connectionName) {
        await context.adapter.signOutUser(csontext, connectionName);
        const adaptiveCard = AdaptiveCardHelper.createSignOutCard();
        return CardResponseHelpers.toSignOutResponse(adaptiveCard);
    }
}
exports.AuthHelper = AuthHelper;