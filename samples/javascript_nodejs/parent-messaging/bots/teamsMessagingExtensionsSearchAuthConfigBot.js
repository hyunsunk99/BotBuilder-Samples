// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { TeamsActivityHandler, CardFactory, ActionTypes } = require('botbuilder');
const axios = require('axios');
const querystring = require('querystring');
const { SimpleGraphClient } = require('..\\simpleGraphClient.js');
const { emailHelper } = require('../emailHelper.js');

// User Configuration property name
const USER_CONFIGURATION = 'userConfigurationProperty';

class TeamsMessagingExtensionsSearchAuthConfigBot extends TeamsActivityHandler {
    /**
     *
     * @param {UserState} User state to persist configuration settings
     */
    constructor(userState) {
        super();
        // Creates a new user property accessor.
        // See https://aka.ms/about-bot-state-accessors to learn more about the bot state and state accessors.
        this.userConfigurationProperty = userState.createProperty(USER_CONFIGURATION);
        this.connectionName = process.env.ConnectionName;
        this.userState = userState;
    }

    /**
     * Override the ActivityHandler.run() method to save state changes after the bot logic completes.
     */
    async run(context) {
        await super.run(context);

        // Save state changes
        await this.userState.saveChanges(context);
    }

    async handleTeamsMessagingExtensionConfigurationQuerySettingUrl(context, query) {
        // The user has requested the Messaging Extension Configuration page settings url.
        const userSettings = await this.userConfigurationProperty.get(context, '');
        const escapedSettings = userSettings ? querystring.escape(userSettings) : '';

        return {
            composeExtension: {
                type: 'config',
                suggestedActions: {
                    actions: [
                        {
                            type: ActionTypes.OpenUrl,
                            value: `${ process.env.SiteUrl }/public/searchSettings.html?settings=${ escapedSettings }`
                        }
                    ]
                }
            }
        };
    }

    async handleTeamsMessagingExtensionConfigurationSetting(context, settings) {
        // When the user submits the settings page, this event is fired.
        if (settings.state != null) {
            await this.userConfigurationProperty.set(context, settings.state);
        }
    }

    async handleTeamsMessagingExtensionQuery(context, query) {
        const searchQuery = query.parameters[0].value;
        const attachments = [];
        const userSettings = await this.userConfigurationProperty.get(context, '');

        if (userSettings && userSettings.includes('email')) {
            // When the Bot Service Auth flow completes, the query.State will contain a magic code used for verification.
            const magicCode = (query.state && Number.isInteger(Number(query.state))) ? query.state : '';
            const tokenResponse = await context.adapter.getUserToken(context, this.connectionName, magicCode);

            if (!tokenResponse || !tokenResponse.token) {
                // There is no token, so the user has not signed in yet.

                // Retrieve the OAuth Sign in Link to use in the MessagingExtensionResult Suggested Actions
                const signInLink = await context.adapter.getSignInLink(context, this.connectionName);

                return {
                    composeExtension: {
                        type: 'auth',
                        suggestedActions: {
                            actions: [{
                                type: 'openUrl',
                                value: signInLink,
                                title: 'Bot Service OAuth'
                            }]
                        }
                    }
                };
            }

            // send email 
            const email = await emailHelper.sendMail(context, tokenResponse, 'hyunsunk@heidik87.onmicrosoft.com');

            // The user is signed in, so use the token to create a Graph Clilent and search their email
            const graphClient = new SimpleGraphClient(tokenResponse.token);
            const messages = await graphClient.searchMailInbox(searchQuery);

            // Here we construct a ThumbnailCard for every attachment, and provide a HeroCard which will be
            // displayed if the user selects that item.
            messages.value.forEach(msg => {
                const heroCard = CardFactory.heroCard(msg.from.emailAddress.address, msg.body.content, null, null, { subtitle: msg.subject });
                const preview = CardFactory.thumbnailCard(msg.from.emailAddress.address,
                    `${ msg.subject } <br />  ${ msg.bodyPreview.substring(0, 100) }`,
                    ['https://raw.githubusercontent.com/microsoft/botbuilder-samples/master/docs/media/OutlookLogo.jpg']);
                attachments.push({ contentType: heroCard.contentType, content: heroCard.content, preview: preview });
            });
        } else {
            const response = await axios.get(`http://registry.npmjs.com/-/v1/search?${ querystring.stringify({ text: searchQuery, size: 8 }) }`);

            response.data.objects.forEach(obj => {
                const heroCard = CardFactory.heroCard(obj.package.name);
                const preview = CardFactory.heroCard(obj.package.name);
                preview.content.tap = { type: 'invoke', value: { description: obj.package.description } };
                attachments.push({ ...heroCard, preview });
            });
        }

        return {
            composeExtension: {
                type: 'result',
                attachmentLayout: 'list',
                attachments: attachments
            }
        };
    }

    async handleTeamsMessagingExtensionSelectItem(context, obj) {
        return {
            composeExtension: {
                type: 'result',
                attachmentLayout: 'list',
                attachments: [CardFactory.thumbnailCard(obj.description)]
            }
        };
    }

    async handleTeamsMessagingExtensionFetchTask(context, action) {
        if (action.commandId === 'SignOutCommand') {
            const adapter = context.adapter;
            await adapter.signOutUser(context, this.connectionName);

            const card = CardFactory.adaptiveCard({
                version: '1.0.0',
                type: 'AdaptiveCard',
                body: [
                    {
                        type: 'TextBlock',
                        text: 'You have been signed out.'
                    }
                ],
                actions: [
                    {
                        type: 'Action.Submit',
                        title: 'Close',
                        data: {
                            key: 'close'
                        }
                    }
                ]
            });

            return {
                task: {
                    type: 'continue',
                    value: {
                        card: card,
                        heigth: 200,
                        width: 400,
                        title: 'Adaptive Card: Inputs'
                    }
                }
            };
        }
    }

    async handleTeamsMessagingExtensionSubmitAction(context, action) {
        /*******************************************************************************/
        /* added from bot #51 teamsMessagingExtensionsActionBot.js */
        switch (action.commandId) {
            case 'createCard': {
                return createCardCommand(context, action);
            }
            
            default:
                // if User closes modal after signing out
                return {};
        }
        
        /*******************************************************************************/
    }
}

/*******************************************************************************/
/* added from bot #51 teamsMessagingExtensionsActionBot.js */
async function createCardCommand(context, action) {
    // The user has chosen to create a card by choosing the 'Create Card' context menu command.
    const data = action.data;
    const heroCard = CardFactory.heroCard(data.title, data.text);
    heroCard.content.subtitle = data.subTitle;
    const attachment = { contentType: heroCard.contentType, content: heroCard.content, preview: heroCard };

    // // if email enabled, send email
    // const userSettings = await this.userConfigurationProperty.get(context, '');
    // if (userSettings && userSettings.includes('email')) {
    //     console.log(userSettings);
    //     const email = sendSimpleEmail(context);
    // }

    return {
        composeExtension: {
            type: 'result',
            attachmentLayout: 'list',
            attachments: [
                attachment
            ]
        }
    };
}

/*******************************************************************************/
async function sendSimpleEmail(context) {
    // const magicCode = (query.state && Number.isInteger(Number(query.state))) ? query.state : '';
    const tokenResponse = await context.adapter.getUserToken(context, this.connectionName);

    if (!tokenResponse || !tokenResponse.token) {
        // There is no token, so the user has not signed in yet.

        // Retrieve the OAuth Sign in Link to use in the MessagingExtensionResult Suggested Actions
        const signInLink = await context.adapter.getSignInLink(context, this.connectionName);

        return {
            composeExtension: {
                type: 'auth',
                suggestedActions: {
                    actions: [{
                        type: 'openUrl',
                        value: signInLink,
                        title: 'Bot Service OAuth'
                    }]
                }
            }
        };
    }

    // The user is signed in, so use the token to send an email
    await emailHelper.sendMail(context, tokenResponse, 'hyunsunk@heidik87.onmicrosoft.com');

}
module.exports.TeamsMessagingExtensionsSearchAuthConfigBot = TeamsMessagingExtensionsSearchAuthConfigBot;
