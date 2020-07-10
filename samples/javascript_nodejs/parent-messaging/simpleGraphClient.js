// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { Client } = require('@microsoft/microsoft-graph-client');

/**
 * This class is a wrapper for the Microsoft Graph API.
 * See: https://developer.microsoft.com/en-us/graph for more information.
 */
class SimpleGraphClient {
    constructor(token) {
        if (!token || !token.trim()) {
            throw new Error('SimpleGraphClient: Invalid token received.');
        }

        this._token = token;

        // Get an Authenticated Microsoft Graph client using the token issued to the user.
        this.graphClient = Client.init({
            authProvider: (done) => {
                done(null, this._token); // First parameter takes an error if you can't get an access token.
            }
        });
    }

    /**
     *
     * @param {string} searchQuery text to search the inbox for.
     */
    async searchMailInbox(searchQuery) {
        // Searches the user's mail Inbox using the Microsoft Graph API
        return await this.graphClient.api('me/mailfolders/inbox/messages').search(searchQuery).get();
    }

    /**************************************************************************/
    // PULLED FROM FORMER BOT #24 CODE
    /**	
     * Sends an email on the user's behalf.	
     * @param {string} toAddress Email address of the email's recipient.	
     * @param {string} subject Subject of the email to be sent to the recipient.	
     * @param {string} content Email message to be sent to the recipient.	
     */	
    async sendMail(toAddress, subject, content) {	
        if (!toAddress || !toAddress.trim()) {	
            throw new Error('SimpleGraphClient.sendMail(): Invalid `toAddress` parameter received.');	
        }	
        if (!subject || !subject.trim()) {	
            throw new Error('SimpleGraphClient.sendMail(): Invalid `subject`  parameter received.');	
        }	
        if (!content || !content.trim()) {	
            throw new Error('SimpleGraphClient.sendMail(): Invalid `content` parameter received.');	
        }	

        // Create the email.	
        const mail = {	
            body: {	
                content: content, // `Hi there! I had this message sent from a bot. - Your friend, ${ graphData.displayName }!`,	
                contentType: 'Text'	
            },	
            subject: subject, // `Message from a bot!`,	
            toRecipients: [{	
                emailAddress: {	
                    address: toAddress	
                }	
            }]	
        };	

        // Send the message.	
        return await this.graphClient	
            .api('/me/sendMail')	
            .post({ message: mail }, (error, res) => {	
                if (error) {	
                    throw error;	
                } else {	
                    return res;	
                }	
            });	
    }	

    /**	
     * Gets recent mail the user has received within the last hour and displays up to 5 of the emails in the bot.	
     */	
    async getRecentMail() {	
        return await this.graphClient	
            .api('/me/messages')	
            .version('beta')	
            .top(5)	
            .get().then((res) => {	
                return res;	
            });	
    }
    
    /**
     * Collects information about the user in the bot.
     */
    async getMe() {
        return await this.graphClient
            .api('/me')
            .get().then((res) => {
                return res;
            });
    }
}

exports.SimpleGraphClient = SimpleGraphClient;
