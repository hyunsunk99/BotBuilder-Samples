// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { TeamsActivityHandler, TeamsInfo, MessageFactory } = require('botbuilder');
const { AuthHelper } = require('../authHelper.js');
const { AdaptiveCardHelper } = require('../adaptiveCardHelper.js');
const { CardResponseHelpers } = require('../cardResponseHelpers.js');
const { EmailHelper } = require('../emailHelper.js');

class TeamsFamilyMessagingBot extends TeamsActivityHandler {
    constructor() {
        super();
        // No longer using Twilio API due to data privacy concerns with 3rd party services
        // this.twilioAccountSid = process.env.TwilioAccountSid;
        // this.twilioAuthToken = process.env.TwilioAuthToken;
        // this.TwilioClient = require('twilio')(this.twilioAccountSid, this.twilioAuthToken);
    };

    async handleTeamsMessagingExtensionFetchTask(context, action) {
        switch(action.commandId)
        {
            // SSO-enabled email through embedded OWA modal
            case "EmailAuthCommand": {
                return await AuthHelper.handleEmailAuthCommand(context);
            }

            // Currently only an Adaptive Card-based demo for Teams-to-TFL UI (no endpoints)
            case "TeamsCommand": {
                return await AuthHelper.handleTeamsCommand(context);
            }

            // NOTE: privacy concerns w/ 3rd party collaborations means 
            // production-level SMS is on the backburner until MS resource exists 
            // case "TextCommand": {              
            // }
            
            default: 
                return null;
        }
    }

    async handleTeamsMessagingExtensionSubmitAction(context, action) {
        // User has submitted template
        console.log(action.commandId);

        // Email 
        if (action.commandId === "EmailAuthCommand") {
            return {};
        }

        // Teams-TFL  
        const token = await AuthHelper.getUserToken(context);
        
        const senderEmail = await EmailHelper.listEmailAddress(context, token);

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

        // Form recipient group ID
        const recipientGroupID = teamName+' Parents & Guardians';

        // Message template submitted
        const submittedData = action.data;

        const adaptiveCard = AdaptiveCardHelper.createAdaptiveCardAttachment(submittedData, senderEmail, recipientGroupID);
        // Display submitted data for preview 
        return CardResponseHelpers.toMessagingExtensionBotMessagePreviewResponse(adaptiveCard);
    }

    handleTeamsMessagingExtensionBotMessagePreviewEdit(context, action) {
        // User chose to edit from preview
        // console.log('previewEdit');
        // The data has been returned to the bot in the action structure.
        const submitData = AdaptiveCardHelper.toSubmitExampleData(action);
        console.log(submitData);

        // var contactEmails = [];
        const selectedRecipients = submitData.SelectedRecipients ? submitData.SelectedRecipients:submitData.AllRecipients;

        const allContactsFormatted = JSON.parse(submitData.AllRecipients);

        console.log('data submitted');
        // This is a preview edit call and so this time we want to re-create the adaptive card editor.
        const adaptiveCard = AdaptiveCardHelper.createAdaptiveCardEditor(submitData.SenderEmail, submitData.RecipientGroupID, allContactsFormatted, selectedRecipients, submitData.Subject, submitData.Body, submitData.SendToChat);
        console.log('card created');
        return CardResponseHelpers.toTaskModuleResponse(adaptiveCard);
    }

    async handleTeamsMessagingExtensionBotMessagePreviewSend(context, action) {
        // console.log('previewSend');
        // The data has been returned to the bot in the action structure.
        const submitData = AdaptiveCardHelper.toSubmitExampleData(action);

        // console.log(submitData);

        if (action.commandId ==='EmailAuthCommand') {
            // User has submitted + confirmed email
            const token = await AuthHelper.getUserToken(context);

            // console.log(submitData.SelectedRecipients);
            await EmailHelper.sendMailToParentsAndGuardians(context, token, submitData.SelectedRecipients, submitData.Subject, submitData.Body);

            return {};
        }
    }

    async handleTeamsMessagingExtensionCardButtonClicked(context, obj) {
        // If the adaptive card was added to the compose window (by either the handleTeamsMessagingExtensionSubmitAction or
        // handleTeamsMessagingExtensionBotMessagePreviewSend handler's return values) the submit values will come in here.
        // console.log('answered');
        const reply = MessageFactory.text('handleTeamsMessagingExtensionCardButtonClicked Value: ' + JSON.stringify(context.activity.value));
        await context.sendActivity(reply);
    }
}

module.exports.TeamsFamilyMessagingBot = TeamsFamilyMessagingBot;
