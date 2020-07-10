// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { CardFactory } = require('botbuilder');

class AdaptiveCardHelper {
    static toSubmitExampleData(action) {
        const activityPreview = action.botActivityPreview[0];
        const attachmentContent = activityPreview.attachments[0].content;
        const userText = attachmentContent.body[1].text;
        const choiceSet = attachmentContent.body[3];
        return {
            MultiSelect: choiceSet.isMultiSelect ? 'true' : 'false',
            Option1: choiceSet.choices[0].title,
            Option2: choiceSet.choices[1].title,
            Option3: choiceSet.choices[2].title,
            Question: userText
        };
    }

    static createAdaptiveCardEditor(userText = null, isMultiSelect = true, option1 = null, option2 = null, option3 = null) {
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
                    text: "Here's an Email Template",
                    type: 'TextBlock',
                    weight: 'bolder'
                },
                { type: 'TextBlock', text: 'Check Email Recipients:' },
                {
                    id: 'Question',
                    placeholder: 'AP Biology Parents (filter by expanding)',
                    type: 'Input.Text',
                    value: userText
                },
                { type: 'TextBlock', text: 'Email Subject:' },
                {
                    id: 'Question',
                    placeholder: 'e.g. Field Trip Reminder',
                    type: 'Input.Text',
                    value: userText
                },
                { type: 'TextBlock', text: 'Send to class chat?' },
                {
                    choices: [{ title: 'Yes', value: 'true' }, { title: 'No', value: 'false' }],
                    id: 'MultiSelect',
                    isMultiSelect: false,
                    style: 'expanded',
                    type: 'Input.ChoiceSet',
                    value: isMultiSelect ? 'true' : 'false'
                },
                { type: 'TextBlock', text: 'Email body' },
                {
                    id: 'Question',
                    placeholder: "e.g. Please remember to sign your child's permission slip for the field trip to the zoo.",
                    type: 'Input.Text',
                    value: userText
                }
            ],
            type: 'AdaptiveCard',
            version: '1.0'
        });
    }

    static createAdaptiveCardAttachment(data) {
        return CardFactory.adaptiveCard({
            actions: [
                { type: 'Action.Submit', title: 'Submit', data: { submitLocation: 'messagingExtensionSubmit' } }
            ],
            body: [
                { text: 'Adaptive Card from Task Module', type: 'TextBlock', weight: 'bolder' },
                { text: `${ data.Question }`, type: 'TextBlock', id: 'Question' },
                { id: 'Answer', placeholder: 'Answer here...', type: 'Input.Text' },
                {
                    choices: [
                        { title: data.Option1, value: data.Option1 },
                        { title: data.Option2, value: data.Option2 },
                        { title: data.Option3, value: data.Option3 }
                    ],
                    id: 'Choices',
                    isMultiSelect: data.MultiSelect,
                    style: 'expanded',
                    type: 'Input.ChoiceSet'
                }
            ],
            type: 'AdaptiveCard',
            version: '1.0'
        });
    }
}
exports.AdaptiveCardHelper = AdaptiveCardHelper;
