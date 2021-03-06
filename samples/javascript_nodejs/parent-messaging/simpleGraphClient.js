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
            },
            defaultVersion: 'beta'
        });
    }

    /**************************************************************************/
    // PULLED FROM FORMER BOT #24 CODE
    /**	
     * Sends an email on the user's behalf.	
     * @param {string} toAddress Email address of the email's recipient.	
     * @param {string} subject Subject of the email to be sent to the recipient.	
     * @param {string} content Email message to be sent to the recipient.	
     */	
    async sendMail(userID='',toAddress, subject, content) {	
        if (!toAddress || !toAddress.trim()) {	
            throw new Error(`SimpleGraphClient.sendMail(): Invalid toAddress parameter ${toAddress} received.`);	
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
                content: content, 	
                contentType: 'Text'	
            },	
            subject: subject, 	
            toRecipients: [{	
                emailAddress: {	
                    address: toAddress	
                }	
            }]	
        };	

        // Send the message.	
        return await this.graphClient	
            .api(`/users/${userID}/sendMail`)	
            .post({ message: mail }, (error, res) => {	
                if (error) {	
                    throw error;	
                } else {	
                    return res;	
                }	
            });	
    }	
    
    /**	
     * Drafts an email on the user's behalf.	
     * @param {string} userID ID of signed-in sender	
     * @param {string} toAddresses Email addresses of recipients.	
     * @param {string} ccAddresses Email addresses of CC'ed.	
     * @param {string} bccAddresses Email addresses of BCC'ed.	
     * @param {string} subject Subject of the email to be sent to the recipient.	
     * @param {string} content Email message to be sent to the recipient.	
     */	
    async draftMessageAndGetID(userID, toAddresses, ccAddresses, bccAddresses, subject='', content='') {	
        if (!userID || !userID.trim()) {	
            throw new Error(`SimpleGraphClient.sendMail(): Invalid userID parameter ${userID} received.`);	
        }	
        // if (!toAddress || !toAddresses.trim()) {	
        //     throw new Error(`SimpleGraphClient.sendMail(): Invalid toAddress parameter ${toAddress} received.`);	
        // }	
        // if (!subject || !subject.trim()) {	
        //     throw new Error('SimpleGraphClient.sendMail(): Invalid `subject` parameter received.');	
        // }	
        // if (!content || !content.trim()) {	
        //     throw new Error('SimpleGraphClient.sendMail(): Invalid `content` parameter received.');	
        // }	

        // Create the email.	
        const message = {	
            body: {	
                content: content, 	
                contentType: 'Text'	
            },	
            subject: subject,
            toRecipients: [],
            ccRecipients: [],
            bccRecipients: [],
        };	

        // Populate direct recipients
        for (let toAddress of toAddresses) {
            console.log(toAddress);
            message["toRecipients"].push(
                {
                    emailAddress: {
                        address: toAddress.emailAddress,
                        name: toAddress.displayName
                    }
                }
            );
        }

        console.log("DRAFTING: ");
        console.dir(message);
        console.dir(message.toRecipients);
        // Populate CC'ed recipients
        for (let ccAddress of ccAddresses) {
            message["ccRecipients"].push(
                {
                    emailAddress: {
                        address: ccAddress.emailAddress,
                        name: ccAddress.displayName
                    }
                }
            );
        }

        // Populate BCC'ed recipients
        for (let bccAddress of bccAddresses) {
            message["bccRecipients"].push(
                {
                    emailAddress: {
                        address: bccAddress.emailAddress,
                        name: bccAddress.displayNameaddress
                    }
                }
            );
        }

        // Post the message.	
        return await this.graphClient.api(`/users/${userID}/messages`)	
            .post( message)
            .then((data) => data.id);
    }	

    /**
     * Collects information about the user in the bot.
     */
    async getUser(userID='') {
        console.log("getUser call");
        console.log("User ID: " + userID);
        return await this.graphClient
            .api(`/users/${userID}`)
            .get().then((res) => {
                return res;
        }).catch(function (error) {
            console.log(error);
        });
    }

    /* Get EduRoster data */
    async getEduRoster(classID='') {
        return await this.graphClient.api(`/education/classes/${classID}/members?$select=id,primaryRole`)
        .get()
        .catch(function (error) {
            console.log(error);
        });
    }

    /* Get user data for a given student ID*/
    async getStudentData(studentID='') {
        return await this.graphClient.api(`/education/users/${studentID}`)
        .get();
    }

    /* Get related contacts and id for a given student */
    /* select statement does not allow exclusive return of relatedContacts */
    async getRelatedContactsAndId(studentID='') {
        return await this.graphClient.api(`/education/users/${studentID}?select=id,relatedContacts`)
        .get();
    }
}

exports.SimpleGraphClient = SimpleGraphClient;
