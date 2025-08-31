require('dotenv').config();
const restify = require('restify');
const axios = require('axios');
const {
  ActivityHandler,
  CloudAdapter,
  ConfigurationServiceClientCredentialFactory,
  createBotFrameworkAuthenticationFromConfiguration,
} = require('botbuilder');

/** ---- Flowise call helper (tolerant to different payloads) ---- */
async function askFlowise(question) {
  if (!process.env.FLOWISE_URL) {
    return '[Flowise not configured: set FLOWISE_URL in .env]';
  }

  const headers = { 'Content-Type': 'application/json' };
  if (process.env.FLOWISE_API_KEY) {
    headers['Authorization'] = `Bearer ${process.env.FLOWISE_API_KEY}`;
    headers['x-api-key'] = process.env.FLOWISE_API_KEY;
  }

  try {
    const res = await axios.post(
      process.env.FLOWISE_URL,
      { question },
      { headers, timeout: 60_000 }
    );

    const d = res.data;
    if (!d) return '[No data from Flowise]';
    if (typeof d === 'string') return d;
    if (d.text) return d.text;
    if (d.answer) return d.answer;
    if (d.result) return d.result;
    if (Array.isArray(d) && d.length && (d[0].text || d[0].answer)) {
      return d[0].text || d[0].answer;
    }
    if (d.data && (d.data.text || d.data.answer)) {
      return d.data.text || d.data.answer;
    }
    return '`[Unrecognized Flowise response]` ' + '```json\n' + JSON.stringify(d, null, 2) + '\n```';
  } catch (err) {
    console.error('Flowise error:', err.response?.status, err.response?.data || err.message);
    if (err.response) {
      return `[Flowise error ${err.response.status}]`;
    }
    return `[Flowise error: ${err.message}]`;
  }
}

/** ---- Bot ---- */
class FlowiseBot extends ActivityHandler {
  constructor() {
    super();

    this.onMessage(async (context, next) => {
      // Extract plain text and strip @Bot mentions if present
      let text = (context.activity && context.activity.text) || '';
      if (context.activity.entities) {
        for (const entity of context.activity.entities) {
          if (
            entity.type === 'mention' &&
            entity.mentioned &&
            entity.mentioned.id === context.activity.recipient.id
          ) {
            text = text.replace(entity.text, '').trim();
          }
        }
      }

      if (!text.trim()) {
        await context.sendActivity('Please send a message.');
        return;
      }

      await context.sendActivity({ type: 'typing' });
      const answer = await askFlowise(text);

      // In channels, reply in-thread
      const reply = { type: 'message', text: answer };
      if (context.activity.conversation.conversationType === 'channel') {
        reply.id = context.activity.id; // thread it
      }

      await context.sendActivity(reply);
      await next();
    });

    this.onMembersAdded(async (context, next) => {
      const membersAdded = context.activity.membersAdded || [];
      for (const member of membersAdded) {
        if (member.id !== context.activity.recipient.id) {
          await context.sendActivity(
            'Hi! I’m connected to Flowise. Mention me with @BotName in a channel or message me directly to ask questions.'
          );
        }
      }
      await next();
    });
  }
}

/** Auth + adapter as before */
const credentialsFactory = new ConfigurationServiceClientCredentialFactory({
  MicrosoftAppId: process.env.MICROSOFT_APP_ID,
  MicrosoftAppPassword: process.env.MICROSOFT_APP_PASSWORD,
  MicrosoftAppType: 'SingleTenant',
  MicrosoftAppTenantId: process.env.MICROSOFT_APP_TENANT_ID,
});
const botFrameworkAuthentication =
  createBotFrameworkAuthenticationFromConfiguration(null, credentialsFactory);
const adapter = new CloudAdapter(botFrameworkAuthentication);

adapter.onTurnError = async (context, error) => {
  console.error('onTurnError:', error);
  await context.sendActivity('Oops. Something went wrong.');
};

const bot = new FlowiseBot();

/** Server */
const server = restify.createServer();
server.use(restify.plugins.bodyParser());

const port = process.env.PORT || 3978;
server.listen(port, () => {
  console.log(`✅ Server listening on port ${port}`);
});

server.post('/api/messages', async (req, res) => {
  await adapter.process(req, res, (context) => bot.run(context));
});

server.get('/', (req, res, next) => {
  res.send(200, { status: 'ok' });
  next();
});
