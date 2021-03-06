// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.
const { TeamsInfo } = require('botbuilder');
const { SimpleGraphClient } = require('./simpleGraphClient.js');

/**
 * These methods call the Microsoft Graph API. The following OAuth scopes are used:
 * 'OpenId' 'email' 'Mail.Send.Shared' 'Mail.Read' 'profile' 'User.Read' 'User.ReadBasic.All'
 * for more information about scopes see:
 * https://developer.microsoft.com/en-us/graph/docs/concepts/permissions_reference
 */
class EmailHelper {
    /**
     * Enable the user to send an email via the bot.
     * @param {TurnContext} context A TurnContext instance containing all the data needed for processing this conversation turn.	
     * @param {Token} token A user token.	
     * @param {string} emailAddress The email address of the recipient.	
     */	
    static async sendMailToParentsAndGuardians(context, token, recipientString, subject, body) {	
        if (!context) {	
            throw new Error('EmailHelper.sendMailToParentsAndGuardians(): `context` cannot be undefined.');	
        }	
        if (!token) {	
            throw new Error('EmailHelper.sendMailToParentsAndGuardians(): `token` cannot be undefined.');	
        }	

        // AAD object id of current user
        const userID = context.activity.from.aadObjectId;

        const contactEmails = recipientString.split(",");
        // console.log(contactEmails);

        const client = new SimpleGraphClient(token);

        // Loop through parent/guardian emails and send email to each
        for (let contactEmail of contactEmails) {
            await client.sendMail(
                userID, 
                contactEmail,	
                `${subject}`,	
                `${body}`	
            );	
        }
    }	

    /**	
     * Send the user their Graph Display Name from the bot.
     * @param {TurnContext} context A TurnContext instance containing all the data needed for processing this conversation turn.
     * @param {Token} token A user token.
     */
    static async listMe(context, token) {
        if (!context) {
            throw new Error('OAuthHelpers.listMe(): `context` cannot be undefined.');
        }
        if (!token) {
            throw new Error('OAuthHelpers.listMe(): `token` cannot be undefined.');
        }

        // Pull in the data from Microsoft Graph.
        const client = new SimpleGraphClient(token);
        const me = await client.getMe();

        return `${me.displayName}`;
    }

    /**
     * Send the user their Graph Email Address from the bot.
     * @param {TurnContext} context A TurnContext instance containing all the data needed for processing this conversation turn.
     * @param {Tokene} token A user token.
     */
    static async listEmailAddress(context, token) {
        if (!context) {
            throw new Error('OAuthHelpers.listEmailAddress(): `context` cannot be undefined.');
        }
        if (!token) {
            throw new Error('OAuthHelpers.listEmailAddress(): `token` cannot be undefined.');
        }

        // AAD object id of current user
        const userID = context.activity.from.aadObjectId;
        
        // Pull in the data from Microsoft Graph.
        const client = new SimpleGraphClient(token);
        const user = await client.getUser(userID);

        return `${user.mail}`;
    }
}

exports.EmailHelper = EmailHelper;
