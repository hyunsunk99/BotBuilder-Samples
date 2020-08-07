// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { MessageFactory, InputHints } = require('botbuilder');

class CardResponseHelpers {
    static toTaskModuleResponse(cardAttachment) {
        // replace card w/ url 
        const draftId = encodeURIComponent("AAMkADY1YmVmY2I4LWVmMzQtNDUzMi1hNjg1LTRiZjI3MjY0NWZjNQBGAAAAAACzJIP-4jG6Qo3BvDLFznWABwDZNhZqXLh7R5QRe-_fqo6YAAAAAAEPAADZNhZqXLh7R5QRe-_fqo6YAAAeQR4IAAA=");
        console.log(draftId);
        const outlookOrigin = 'https://outlook.office.com';
        // need approval for EDU app domain if hosted there 
        // var src = `${outlookOrigin}/mail/opxdeeplink/compose/${draftId}?isanonymous=true&opxAuth&hostApp=teams&useOwaTheme`;

        var src = `${outlookOrigin}/mail/opxdeeplink/compose/${draftId}?isanonymous=true&opxAuth&hostApp=teams&useOwaTheme&cspoff`;
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
    // static toTaskModuleResponse(cardAttachment) {
    //     return {
    //         task: {
    //             type: 'continue',
    //             value: {
    //                 card: cardAttachment,
    //                 height: 500,
    //                 title: 'Email Editor',
    //                 url: null,
    //                 width: 800
    //             }
    //         }
    //     };
    // }

    static toMessagingExtensionBotMessagePreviewResponse(cardAttachment) {
        return {
            composeExtension: {
                activityPreview: MessageFactory.attachment(cardAttachment, null, null, InputHints.ExpectingInput),
                type: 'botMessagePreview'
            }
        };
    }

    static toEmailSentResponse(cardAttachment) {
        return {
            task: {
                type: 'continue',
                value: {
                    card: cardAttachment,
                    heigth: 200,
                    width: 400,
                    title: 'Adaptive Card: Inputs'
                }
            }
        };
    }
}

exports.CardResponseHelpers = CardResponseHelpers;
