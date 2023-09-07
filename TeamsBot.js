const { TeamsActivityHandler, TurnContext, CardFactory, TeamsInfo } = require("botbuilder");
const fs = require('fs')

function createCardCommand(context, action) {
  // The user has chosen to create a card by choosing the 'Create Card' context menu command.
  const data = action.data;
  const heroCard = CardFactory.heroCard(`Nishant 007 ${data.test} ${action.messagePayload.body.content}`);
  heroCard.content.subtitle = data.subTitle;
  const attachment = { contentType: heroCard.contentType, content: heroCard.content, preview: heroCard };

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

const getSampleAdaptiveCard = (name) => ({
    "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
    "type": "AdaptiveCard",
    "version": "1.6",
    "minHeight": "100px",
    "body": [
        {
            "type": "TextBlock",
            "wrap": true,
            "text": `Welcome ${name}`
        }
    ],
    "actions": [
        {
            "type": "Action.Submit",
            "title": "Thanks",
            "value": "clicked"
        }
    ]
});
class TeamsBot extends TeamsActivityHandler {
  constructor(conversationReferences) {
    super();
    // Dependency injected dictionary for storing ConversationReference objects used in NotifyController to proactively message users
    
    this.conversationReferences = conversationReferences;
    

    this.onConversationUpdate(async (context, next) => {
      this.addConversationReference(context.activity);

      await next();
    });

    this.onMembersAdded(async (context, next) => {
      const membersAdded = context.activity.membersAdded;
      for (let cnt = 0; cnt < membersAdded.length; cnt++) {
        if (membersAdded[cnt].id !== context.activity.recipient.id) {
        //   const welcomeMessage = `Welcome ${membersAdded[cnt].name}`;
          console.log(membersAdded[cnt]);
          const card = CardFactory.adaptiveCard(getSampleAdaptiveCard(membersAdded[cnt].aadObjectId));
          await context.sendActivity({ attachments: [card] });
        }
      }

      // By calling next() you ensure that the next BotHandler is run.
      await next();
    });

    this.onMessage(async (context, next) => {
        console.log("received message", context.activity);
        this.addConversationReference(context.activity);
        
        const activity = context.activity;
        const connector = context.adapter.createConnectorClient(activity.serviceUrl);
        const response = await connector.conversations.getConversationMembers(activity.conversation.id);
        const convMemJson = JSON.stringify(response);

        const channels = await TeamsInfo.getTeamChannels(context);
        const channelsJson = JSON.stringify(channels);

        // Echo back what the user said
        await context.sendActivity(`You sent '${context.activity.text}'. Navigate to http://{your-domain}/api/notify to proactively message everyone who has previously messaged this bot`);
        
        await context.sendActivity(`channelsJson: ${channelsJson}`);
        await context.sendActivity(`convMemJson: ${convMemJson}`);
        console.log("Send the reply");
        await next();
    });

  }

  async handleTeamsMessagingExtensionSubmitAction(context, action) {
    console.log(action);
    if(action.commandId === 'test') {
      return createCardCommand(context, action);
    } else {
      return createCardCommand(context, action);
    }
  }

  addConversationReference(activity) {
    const conversationReference = TurnContext.getConversationReference(activity);
    this.conversationReferences[conversationReference.conversation.id] = conversationReference;
    fs.writeFileSync('conv.json', JSON.stringify(this.conversationReferences))
  }

}

module.exports.TeamsBot = TeamsBot;