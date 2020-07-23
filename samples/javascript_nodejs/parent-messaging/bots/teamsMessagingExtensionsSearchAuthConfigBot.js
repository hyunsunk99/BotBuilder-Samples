// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { TeamsActivityHandler, TeamsInfo, MessageFactory } = require('botbuilder');
const { AuthHelper } = require('../authHelper.js');
const { AdaptiveCardHelper } = require('../adaptiveCardHelper.js');
const { CardResponseHelpers } = require('../cardResponseHelpers.js');
const { SimpleGraphClient } = require('../simpleGraphClient.js');

// Removed user configuration settings (possibly add back for preferred mode of comm [Email, SMS, etc.])
// User Configuration property name
// const USER_CONFIGURATION = 'userConfigurationProperty';

class TeamsMessagingExtensionsSearchAuthConfigBot extends TeamsActivityHandler {
    constructor() {
        super();
        this.connectionName = process.env.ConnectionName;
    };

    // REMOVED USER CONFIG SETTINGS (consider re-implementing later)
    // UPDATE canUpdateConfiguration in manifest.json if re-implementing
    // /**
    //  * Override the ActivityHandler.run() method to save state changes after the bot logic completes.
    //  */
    // async run(context) {
    //     await super.run(context);

    //     // Save state changes
    //     await this.userState.saveChanges(context);
    // }

    // async handleTeamsMessagingExtensionConfigurationQuerySettingUrl(context, query) {
    //     // The user has requested the Messaging Extension Configuration page settings url.
    //     const userSettings = await this.userConfigurationProperty.get(context, '');
    //     const escapedSettings = userSettings ? querystring.escape(userSettings) : '';

    //     return {
    //         composeExtension: {
    //             type: 'config',
    //             suggestedActions: {
    //                 actions: [
    //                     {
    //                         type: ActionTypes.OpenUrl,
    //                         value: `${ process.env.SiteUrl }/public/searchSettings.html?settings=${ escapedSettings }`
    //                     }
    //                 ]
    //             }
    //         }
    //     };
    // }

    // async handleTeamsMessagingExtensionConfigurationSetting(context, settings) {
    //     // When the user submits the settings page, this event is fired.
    //     if (settings.state != null) {
    //         await this.userConfigurationProperty.set(context, settings.state);
    //     }
    
    
    // }

    async handleTeamsMessagingExtensionFetchTask(context, action) {
        switch(action.commandId)
        {
            // Command chosen by user
            case "EmailAuthCommand": {
                return await AuthHelper.handleEmailAuthCommand(context,action,this.connectionName);
            }
            
            case "SignOutCommand":
                return await AuthHelper.handleSignOutCommand(context, this.connectionName);
            
            default: 
                return null;
        }
    }

    handleTeamsMessagingExtensionSubmitAction(context, action) {
        console.log('submitted');
        console.log(action.commandId);
        if (action.commandId==="SignOutCommand") {
            // Close Sign Out confirmation modal 
            return {};
        }

        // Message template submitted
        const submittedData = action.data;

        // TEMPORARY SENDER EMAIL AND RECIPIENT GROUP
        // MODIFY TO HANDLE USER MODS FOR RECIPIENTS (filtering, adding)
        const senderEmail = 'hyunsunk@heidik87.onmicrosoft.com'
        // Retrieve recipient group ID (hard code for now)
        const recipientGroupID = 'AP Bio Parents (click to expand)'

        const adaptiveCard = AdaptiveCardHelper.createAdaptiveCardAttachment(submittedData, senderEmail, recipientGroupID);
        // Display submitted data for preview 
        return CardResponseHelpers.toMessagingExtensionBotMessagePreviewResponse(adaptiveCard);
    }

    handleTeamsMessagingExtensionBotMessagePreviewEdit(context, action) {
        // User chose to edit from preview
        console.log('previewEdit');
        // The data has been returned to the bot in the action structure.
        const submitData = AdaptiveCardHelper.toSubmitExampleData(action);

        console.log('data submitted');
        // This is a preview edit call and so this time we want to re-create the adaptive card editor.
        const adaptiveCard = AdaptiveCardHelper.createAdaptiveCardEditor(submitData.SenderEmail, submitData.RecipientGroupID, submitData.Subject, submitData.Body, submitData.SendToChat);
        console.log('card created');
        return CardResponseHelpers.toTaskModuleResponse(adaptiveCard);
    }

    handleTeamsMessagingExtensionBotMessagePreviewSend(context, action) {
        console.log('previewSend');
        // The data has been returned to the bot in the action structure.
        const submitData = AdaptiveCardHelper.toSubmitExampleData(action);

        // ADD CONDITION FOR IFSENDTOSTUDENTS
        // if user selected to send to class channel:
        // This is a send so we are done and we will create the adaptive card editor.
        // const adaptiveCard = AdaptiveCardHelper.createAdaptiveCardAttachment(submitData);
        // const responseActivity = { type: 'message', attachments: [adaptiveCard] };
        // await context.sendActivity(responseActivity);

        if (action.commandId ==='EmailAuthCommand') {
            // User has submitted + confirmed email
            const adaptiveCard = AdaptiveCardHelper.createEmailSentCard();
            return CardResponseHelpers.toSignOutResponse(adaptiveCard);
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
