// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { MessageFactory, InputHints } = require('botbuilder');

class CardResponseHelpers {
    static toEmailCommandResponse(draftID) {

        const encodedDraftID = encodeURIComponent(draftID);
        const outlookOrigin = 'https://outlook.office.com';

        var src = `${outlookOrigin}/mail/opxdeeplink/compose/${encodedDraftID}?isanonymous=true&hostApp=teams&opxAuth&useOwaTheme`;
        
        return {
            task: {
                type: 'continue',
                value: {
                    url: `${src}`,
                    height: 500,
                    title: 'Email Editor',
                    width: 800
                }
            }
        };
    }

    static toTaskModuleResponse(cardAttachment) {
        return {
            task: {
                type: 'continue',
                value: {
                    card: cardAttachment,
                    height: 500,
                    title: 'Teams Card Editor',
                    url: null,
                    width: 800
                }
            }
        };
    }

    static toMessagingExtensionBotMessagePreviewResponse(cardAttachment) {
        return {
            composeExtension: {
                activityPreview: MessageFactory.attachment(cardAttachment, null, null, InputHints.ExpectingInput),
                type: 'botMessagePreview'
            }
        };
    }
}

exports.CardResponseHelpers = CardResponseHelpers;
