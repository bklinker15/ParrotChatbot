// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const {
    TurnContext,
    CardFactory,
    MessageFactory,
    TeamsInfo,
    TeamsActivityHandler,
    ActionTypes
} = require('botbuilder');

class ParrotBot extends TeamsActivityHandler  {
    constructor() {
        super();
        
        this.onTeamsChannelCreatedEvent(async (channelInfo, teamInfo, turnContext, next) => {
            await context.sendActivity('Hello! I am your personal Parrot. You can send me things you want to remember or share to other devices logged in to Teams.');

            await next();
        });
        
        this.onTeamsMembersAddedEvent(async (membersAdded, teamInfo, turnContext, next) => {
            const card = CardFactory.heroCard('Welcome to Parrot!', 'Hello! I am your personal Parrot. You can send me things you want to remember or share to other devices logged in to Teams.');
            const message = MessageFactory.attachment(card);
            await turnContext.sendActivity(message);
            
            await next();
        });
        
        // See https://aka.ms/about-bot-activity-message to learn more about the message and other activity types.
        this.onMessage(async (context, next) => {
            await context.sendActivity(context.activity.text);

            await next();
        });
    }
}

module.exports.ParrotBot = ParrotBot;
