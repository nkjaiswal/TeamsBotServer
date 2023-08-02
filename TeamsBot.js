const { TeamsActivityHandler, TurnContext, CardFactory } = require("botbuilder");
const fs = require('fs')

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

        // Echo back what the user said
        await context.sendActivity(`You sent '${context.activity.text}'. Navigate to http://{your-domain}/api/notify to proactively message everyone who has previously messaged this bot.`);
        console.log("Send the reply");
        await next();
    });

  }

  addConversationReference(activity) {
    const conversationReference = TurnContext.getConversationReference(activity);
    this.conversationReferences[conversationReference.conversation.id] = conversationReference;
    fs.writeFileSync('conv.json', JSON.stringify(this.conversationReferences))
  }

}

module.exports.TeamsBot = TeamsBot;