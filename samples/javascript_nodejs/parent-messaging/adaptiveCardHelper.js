// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { CardFactory } = require('botbuilder');

class AdaptiveCardHelper {
    static toSubmitExampleData(action) {
        const activityPreview = action.botActivityPreview[0];
        const attachmentContent = activityPreview.attachments[0].content;
        const facts = attachmentContent.body[1].facts;
        const subject = attachmentContent.body[2];
        const body = attachmentContent.body[3];
        const sendToChat = attachmentContent.body[4];
        console.log(sendToChat.value);
        return {
            SenderEmail: facts[0].value,
            RecipientGroupID: facts[1].value,
            Subject: subject.text,
            Body: body.text,
            SendToChat: sendToChat.value
        };
    }

    // UPDATE senderEmail with User's verified account
    static createAdaptiveCardEditor(senderEmail=null, recipientGroupID=null, messageSubject=null, messageBody = null, sendToChat = false) {
        return CardFactory.adaptiveCard({
            actions: [
                {
                    data: {
                        submitLocation: 'messagingExtensionFetchTask'
                    },
                    title: 'Submit',
                    type: 'Action.Submit'
                }
            ],
            body: [
                {
                    id: 'Facts',
                    type:'FactSet',
                    separator: true,
                    facts: [
                        { 
                            title: "From:", 
                            value: senderEmail // verified user's email
                        },
                        { 
                            title: "To:", 
                            value: recipientGroupID, // related contacts for current channel members
                        },
                    ]
                },
                { type: 'TextBlock', text: 'Email Subject:' },
                {
                    id: 'Subject',
                    placeholder: 'e.g. Field Trip Rescheduled',
                    type: 'Input.Text',
                    value: messageSubject
                },
                { type: 'TextBlock', text: 'Email body' },
                {
                    id: 'Body',
                    placeholder: "e.g. Our trip to Seattle U's marine biology lab has been postponed until December 1st.",
                    type: 'Input.Text',
                    value: messageBody,
                    isMultiline: true,
                    maxLength: 0,
                    wrap: true
                },
                {
                    title: "Send to class chat",
                    id: 'sendToChat',
                    type: 'Input.Toggle',
                    value: sendToChat,
                },
            ],
            type: 'AdaptiveCard',
            version: '1.0'
        });
    }

    static createAdaptiveCardAttachment(data, senderEmail, recipientGroupID) {
        return CardFactory.adaptiveCard({
            body: [
                { text: "Here's a preview of your message", type: 'TextBlock', weight: 'bolder' },
                {
                    type:'FactSet',
                    separator: true,
                    facts: [
                        { 
                            title: "From:",
                            value: senderEmail
                        },
                        { 
                            title: "To:",
                            value: recipientGroupID
                        },
                    ]
                },
                { text: `${ data.Subject }`, type: 'TextBlock', id: 'Subject', weight: 'bolder' },
                { text: `${ data.Body }`, type: 'TextBlock', id: 'Body', isMultiline: true, maxLength: 0, wrap: true },
                { value: `${ data.sendToChat}`, type: 'Input.Toggle', id: 'SendToChat', isVisible: 'false'}
            ],
            type: 'AdaptiveCard',
            version: '1.0'
        });
    }

    // Simple sign out modal 
    static createSignOutCard() {
        return CardFactory.adaptiveCard({
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
    }

    // Simple email confirmation modal 
    static createEmailSentCard() {
        return CardFactory.adaptiveCard({
            version: '1.0.0',
            type: 'AdaptiveCard',
            body: [
                {
                    type: 'TextBlock',
                    text: 'Your email has been sent.'
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
    }
}
exports.AdaptiveCardHelper = AdaptiveCardHelper;
