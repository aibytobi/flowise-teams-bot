require('dotenv').config();
const restify = require('restify');
const {
  ActivityHandler,
  CloudAdapter,
  ConfigurationServiceClientCredentialFactory,
  createBotFrameworkAuthenticationFromConfiguration,
} = require('botbuilder');

/** Minimal echo bot */
class EchoBot extends ActivityHandler {
  constructor() {
    super();
    this.onMessage(async (context, next) => {
      const text = (context.activity && context.activity.text) || '';
      await context.sendActivity(`echo: ${text}`);
      await next();
    });
    this.onMembersAdded(async (context, next) => {
      const membersAdded = context.activity.membersAdded || [];
      for (const member of membersAdded) {
        if (member.id !== context.activity.recipient.id) {
          await context.sendActivity('Hello! I am alive. Send me a message and I will echo it.');
        }
      }
      await next();
    });
  }
}

/** Bot Framework auth wired to env (we’ll populate these next step) */
const credentialsFactory = new ConfigurationServiceClientCredentialFactory({
  MicrosoftAppId: process.env.MICROSOFT_APP_ID,
  MicrosoftAppPassword: process.env.MICROSOFT_APP_PASSWORD,
  MicrosoftAppType: 'SingleTenant',
  MicrosoftAppTenantId: process.env.MICROSOFT_APP_TENANT_ID,
});
const botFrameworkAuthentication =
  createBotFrameworkAuthenticationFromConfiguration(null, credentialsFactory);
const adapter = new CloudAdapter(botFrameworkAuthentication);

/** Basic error handler */
adapter.onTurnError = async (context, error) => {
  console.error('onTurnError:', error);
  await context.sendActivity('Oops. Something went wrong.');
};

const bot = new EchoBot();

/** Restify server */
const server = restify.createServer();
server.use(restify.plugins.bodyParser());

const port = process.env.PORT || 3978;
server.listen(port, () => {
  console.log(`✅ Server listening on port ${port}`);
});

/** Bot messages endpoint that Azure/Teams will call */
server.post('/api/messages', async (req, res) => {
  await adapter.process(req, res, (context) => bot.run(context));
});

/** Simple health probe */
server.get('/', (req, res, next) => {
  res.send(200, { status: 'ok' });
  next();
});
