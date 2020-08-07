// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { CardFactory } = require('botbuilder');

class AdaptiveCardHelper {
    static toSubmitExampleData(action) {
        const activityPreview = action.botActivityPreview[0];
        const attachmentContent = activityPreview.attachments[0].content;
        console.log('adaptiveCardHelper - SUBMIT EXAMPLE DATA')
        console.dir(attachmentContent);
        const subject = attachmentContent.body[0];
        const facts = attachmentContent.body[1].facts;
        const allRecipients = attachmentContent.body[2]; // array with all recipients in ChoiceSet format
        const selectedRecipients = attachmentContent.body[3].text==='undefined' ? '' : attachmentContent.body[3].text; // string containing selected recipients
        console.log("SELECTED");
        console.log(selectedRecipients)
        const body = attachmentContent.body[4];
        const sendToChat = attachmentContent.body[5];
        return {
            SenderEmail: facts[0].value,
            RecipientGroupID: facts[1].value,
            AllRecipients: allRecipients.text,
            SelectedRecipients: selectedRecipients,
            Subject: subject.text,
            Body: body.text,
            SendToChat: sendToChat.value
        };
    }

    // fields for modal embedded w/ OWA iframe
    // static createAdaptiveCardEditor(senderEmail=null, recipientGroupID=null, contactEmails=[], emailListString='', messageSubject=null, messageBody = null, sendToChat = false) {
    //     // example draft ID for an email currently in this tenant's drafts 
    //     const draftId = "AAMkADY1YmVmY2I4LWVmMzQtNDUzMi1hNjg1LTRiZjI3MjY0NWZjNQBGAAAAAACzJIP-4jG6Qo3BvDLFznWABwDZNhZqXLh7R5QRe-_fqo6YAAAAAAEPAADZNhZqXLh7R5QRe-_fqo6YAAAcqcAQAAA=";
    //     const outlookOrigin = 'https://outlook.office.com';
    // }

    // return adaptive card editor with recipients pre-populated
    static createAdaptiveCardEditor(senderEmail=null, recipientGroupID=null, contactEmails=[], emailListString='', messageSubject=null, messageBody = null, sendToChat = false) {
        return CardFactory.adaptiveCard({
            version: '1.1',
            actions: [
                {
                    data: {
                        submitLocation: 'messagingExtensionFetchTask'
                    },
                    title: 'Preview',
                    type: 'Action.Submit'
                },

            ],
            body: [
                {
                    type: 'ColumnSet',
                    columns: [
                        {
                            type: 'Column',
                            width: 'auto',
                            items: [
                                {
                                    type:'TextBlock',
                                    text: "From: ",
                                    weight: 'bolder'
                                },
                                {
                                    type:'TextBlock',
                                    text: "To: ",
                                    weight: 'bolder'
                                }
                            ]
                        },
                        {
                            type: 'Column',
                            width: 'stretch',
                            items: [
                                {
                                    type:'TextBlock',
                                    text: senderEmail,
                                    horizontalAlignment: 'right'
                                },
                                {
                                    type: "ActionSet",
                                    actions: [
                                        {
                                            type: "Action.ShowCard",
                                            title: recipientGroupID,
                                            card: {
                                                type: "AdaptiveCard",
                                                body: [
                                                    {
                                                        type: 'ColumnSet',
                                                        columns: [
                                                            {
                                                                type: 'Column',
                                                                width: 10,
                                                                items: [
                                                                    {
                                                                        type:'TextBlock',
                                                                        text: ""
                                                                    }
                                                                ]
                                                            },
                                                            {
                                                                type: 'Column',
                                                                width: 15,
                                                                items: [
                                                                    {
                                                                        type: "TextBlock",
                                                                        text: "Include: ",
                                                                    },
                                                                    {
                                                                        type: "Input.ChoiceSet",
                                                                        id: 'recipientList',
                                                                        style: 'expanded',
                                                                        value: emailListString,
                                                                        choices: contactEmails,
                                                                        isMultiSelect: true
                                                                    }
                                                                ]
                                                            },
                                                        ]
                                                    }
                                                ]
                                            }
                                        }
                                    ]
                                }
                            ]
                        }
                    ]
                },
                {
                    id: 'allContacts',
                    type: 'Input.Text',
                    value: JSON.stringify(contactEmails),
                    isVisible: false
                },
                { type: 'TextBlock', text: 'Email Subject:', wrap: true },
                {
                    id: 'subject',
                    placeholder: 'e.g. Field Trip Rescheduled',
                    type: 'Input.Text',
                    value: messageSubject
                },
                { type: 'TextBlock', text: 'Email body' },
                {
                    id: 'body',
                    placeholder: "e.g. Our trip to Seattle U's lab facility has been postponed...",
                    type: 'Input.Text',
                    value: messageBody,
                    isMultiline: true,
                    maxLength: 0,
                    wrap: true,
                    height: 150
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
        // const allContactsFormatted = data.allContacts.split(",");
        console.log("ADAPTIVE CARD");

        // if list was never expanded, recipientList === undefined
        var recipientList = data.recipientList;
        console.log(recipientList);

        // if no recipients were selected (ie user did not click/expand group)
        if (!recipientList) {
            const allContactsParsed = JSON.parse(data.allContacts);
            var allContactsString = '';
            for (let contactInfo of allContactsParsed) {
                allContactsString+=contactInfo.value + ',';
            }
            console.log(allContactsString);
            recipientList=allContactsString;
        }

        return CardFactory.adaptiveCard({
            body: [
                { text: `${ data.subject }`, type: 'TextBlock', id: 'Subject', weight: 'bolder', wrap: true},
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
                { text: `${ data.allContacts }`, type: 'TextBlock', id: 'AllContacts', isVisible: false},
                { text: recipientList, type: 'TextBlock', id: 'RecipientList', isVisible: false},
                { text: `${ data.body }`, type: 'TextBlock', id: 'Body', isMultiline: true, maxLength: 0, wrap: true, separator: true },
                { value: `${ data.sendToChat}`, type: 'Input.Toggle', id: 'SendToChat', isVisible: false}
            ],
            type: 'AdaptiveCard',
            version: '1.0'
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
