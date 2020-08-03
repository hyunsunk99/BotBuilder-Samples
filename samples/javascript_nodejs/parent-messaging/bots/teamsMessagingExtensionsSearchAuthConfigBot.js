// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { TeamsActivityHandler, TeamsInfo, MessageFactory } = require('botbuilder');
const { AuthHelper } = require('../authHelper.js');
const { AdaptiveCardHelper } = require('../adaptiveCardHelper.js');
const { CardResponseHelpers } = require('../cardResponseHelpers.js');
const { EmailHelper } = require('../emailHelper.js');



// Removed user configuration settings (possibly add back for preferred mode of comm [Email, SMS, etc.])
// User Configuration property name
// const USER_CONFIGURATION = 'userConfigurationProperty';

class TeamsMessagingExtensionsSearchAuthConfigBot extends TeamsActivityHandler {
    constructor() {
        super();
        this.twilioAccountSid = process.env.TwilioAccountSid;
        this.twilioAuthToken = process.env.TwilioAuthToken;
        this.TwilioClient = require('twilio')(this.twilioAccountSid, this.twilioAuthToken);
    };

    async handleTeamsMessagingExtensionFetchTask(context, action) {
        switch(action.commandId)
        {
            // Command chosen by user
            case "EmailAuthCommand": {
                return await AuthHelper.handleEmailAuthCommand(context);
            }

            case "TwilioTextCommand": {
                console.log(this.twilioAccountSid);
                await this.TwilioClient.messages
                    .create({
                        body: 'This thing on?',
                        from: '+14044424990',
                        to: '+12015750442'
                    })
                    .then(message => console.log(message.sid));

                console.log("SMS sent");
            }
            
            default: 
                return null;
        }
    }

    async handleTeamsMessagingExtensionSubmitAction(context, action) {
        // User has submitted template
        console.log('submitted');
        console.log(action.commandId);

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
        console.log('previewEdit');
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
        console.log('previewSend');
        // The data has been returned to the bot in the action structure.
        const submitData = AdaptiveCardHelper.toSubmitExampleData(action);

        console.log(submitData);

        // ADD CONDITION FOR IFSENDTOSTUDENTS
        // if user selected to send to class channel:
        // This is a send so we are done and we will create the adaptive card editor.
        // const adaptiveCard = AdaptiveCardHelper.createAdaptiveCardAttachment(submitData);
        // const responseActivity = { type: 'message', attachments: [adaptiveCard] };
        // await context.sendActivity(responseActivity);

        if (action.commandId ==='EmailAuthCommand') {
            // User has submitted + confirmed email
            const token = await AuthHelper.getUserToken(context);

            console.log(submitData.SelectedRecipients);
            await EmailHelper.sendMailToParentsAndGuardians(context, token, submitData.SelectedRecipients, submitData.Subject, submitData.Body);

            /* **** TROUBLESHOOT HERE FOR PROPER MODAL **** */
            const adaptiveCard = AdaptiveCardHelper.createEmailSentCard();
            return CardResponseHelpers.toEmailSentResponse(adaptiveCard);
            // const responseActivity = { type: 'message', attachments: [adaptiveCard] }; //"Your email has been sent"
            // await context.sendActivity(responseActivity);
        }
    }

    async handleTeamsMessagingExtensionCardButtonClicked(context, obj) {
        // If the adaptive card was added to the compose window (by either the handleTeamsMessagingExtensionSubmitAction or
        // handleTeamsMessagingExtensionBotMessagePreviewSend handler's return values) the submit values will come in here.
        console.log('answered');
        const reply = MessageFactory.text('handleTeamsMessagingExtensionCardButtonClicked Value: ' + JSON.stringify(context.activity.value));
        await context.sendActivity(reply);
    }
}

module.exports.TeamsMessagingExtensionsSearchAuthConfigBot = TeamsMessagingExtensionsSearchAuthConfigBot;
