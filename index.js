const {  BotFrameworkAdapter } = require("botbuilder");
const { TeamsBot } = require("./TeamsBot.js");
const express = require('express');
const bodyParser = require('body-parser');
require('dotenv').config();
const app = express()
const port = process.env.PORT || 3000;

app.use(bodyParser.urlencoded({ extended: true }));
app.use(bodyParser.json());
app.use(express.static('./public'));

app.post("/auth-end", (req, res) => {
    const token = req.body.token;
    if(!token) {
      return res.json({});
    }
    const user = JSON.parse(Buffer.from(token.split('.')[1], 'base64').toString());
    res.json(user);
});


const adapter = new BotFrameworkAdapter({
    appId: process.env.MicrosoftAppId,
    appPassword: process.env.MicrosoftAppPassword,
});

const conversationReferences = {};
const bot = new TeamsBot(conversationReferences);

adapter.onTurnError = async (context, error) => {
    console.error(`\n [onTurnError] unhandled error: ${ error }`);
};

app.post("/api/messages", async (req, res) => {
    await adapter.processActivity(req, res, async (context) => {
      await bot.run(context);
    });
});

app.get('/api/notify', async (req, res) => {
    console.log(JSON.stringify(conversationReferences));
    for (const conversationReference of Object.values(conversationReferences)) {
  
      await adapter.continueConversation(conversationReference, async (context) => {
        await context.sendActivity('proactive hello');
      });
    }
  
    res.setHeader('Content-Type', 'text/html');
    res.writeHead(200);
    res.write('<html><body><h1>Proactive messages have been sent.</h1></body></html>');
    res.end();
});

app.listen(port, () => {
  console.log(`Example app listening at http://localhost:${port}`)
})