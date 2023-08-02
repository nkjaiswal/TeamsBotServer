const { TeamsActivityHandler, TurnContext } = require("botbuilder");

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
          const welcomeMessage = 'Welcome to the Proactive Bot sample.  Navigate to http://{your-domain}/api/notify to proactively message everyone who has previously messaged this bot.';
          await context.sendActivity(welcomeMessage);
        }
      }

      // By calling next() you ensure that the next BotHandler is run.
      await next();
    });

    this.onMessage(async (context, next) => {
        console.log("received message", context.activity.text, context.activity.serviceUrl);
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
  }

}

module.exports.TeamsBot = TeamsBot;