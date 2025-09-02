require('dotenv').config();
const restify = require('restify');
const axios = require('axios');
const path = require('path');
const fs = require('fs');

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

/** ---------- Helpers: detect audio attachments from Teams ---------- */
function looksLikeAudioFilename(name = '') {
  const lower = name.toLowerCase();
  return (
    lower.endsWith('.wav') ||
    lower.endsWith('.mp3') ||
    lower.endsWith('.m4a') ||
    lower.endsWith('.ogg') ||
    lower.endsWith('.webm') ||
    lower.endsWith('.aac') ||
    lower.endsWith('.flac')
  );
}

/**
 * Extracts best-guess audio info from a Teams attachment.
 * Supports:
 * - Teams file download card: contentType 'application/vnd.microsoft.teams.file.download.info'
 *   -> use attachment.content.downloadUrl + attachment.name
 * - Direct audio content types: 'audio/*' with contentUrl
 */
function extractAudioInfoFromAttachment(att = {}) {
  const ct = (att.contentType || '').toLowerCase();

  // Case 1: Teams File Download Info card
  if (ct === 'application/vnd.microsoft.teams.file.download.info') {
    const name = att.name || att.content?.name || 'audio';
    const downloadUrl = att.content?.downloadUrl || att.contentUrl;
    const fileType = att.content?.fileType || path.extname(name).replace('.', '');
    if (downloadUrl && looksLikeAudioFilename(name)) {
      return {
        source: 'teams-file-download-info',
        name,
        fileType,
        contentType: ct,
        url: downloadUrl,
      };
    }
  }

  // Case 2: Direct audio content types
  if (ct.startsWith('audio/')) {
    const name = att.name || `audio.${ct.split('/')[1] || 'wav'}`;
    if (att.contentUrl) {
      return {
        source: 'direct-audio',
        name,
        fileType: ct.split('/')[1] || 'wav',
        contentType: ct,
        url: att.contentUrl,
      };
    }
  }

  // Case 3: Unknown type but filename looks like audio
  const nameGuess = att.name || att.content?.name || '';
  if (looksLikeAudioFilename(nameGuess) && (att.content?.downloadUrl || att.contentUrl)) {
    return {
      source: 'filename-audio-fallback',
      name: nameGuess,
      fileType: path.extname(nameGuess).replace('.', '') || 'wav',
      contentType: ct || 'unknown',
      url: att.content?.downloadUrl || att.contentUrl,
    };
  }

  return null;
}

/** ---- Bot ---- */
class FlowiseBot extends ActivityHandler {
  constructor() {
    super();

    this.onMessage(async (context, next) => {
      const activity = context.activity || {};
      const attachments = activity.attachments || [];

      // 1) First, check if the user sent any audio attachment(s)
      const audioItems = [];
      for (const att of attachments) {
        const info = extractAudioInfoFromAttachment(att);
        if (info) audioItems.push(info);
      }

      if (audioItems.length > 0) {
        // For Step 1, we ONLY acknowledge and log the details (no download/transcribe yet)
        for (const a of audioItems) {
          console.log('🎤 Audio attachment detected:', a);
        }

        const list = audioItems
          .map((a, i) => `• ${a.name} (${a.fileType}) via ${a.source}`)
          .join('\n');

        await context.sendActivity(
          `🎧 I received your audio file:\n${list}\n\n(Transcription comes next step.)`
        );

        await next();
        return;
      }

      // 2) Otherwise, fall back to normal text handling (Flowise)
      // Extract plain text and strip @Bot mentions if present
      let text = (activity && activity.text) || '';
      if (activity.entities) {
        for (const entity of activity.entities) {
          if (
            entity.type === 'mention' &&
            entity.mentioned &&
            entity.mentioned.id === activity.recipient.id
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
      if (activity.conversation?.conversationType === 'channel') {
        reply.id = activity.id; // thread it
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
