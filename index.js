require('dotenv').config();
const restify = require('restify');
const axios = require('axios');
const path = require('path');
const fs = require('fs');
const { spawn } = require('child_process');
const ffmpegPath = require('ffmpeg-static');
const sdk = require('microsoft-cognitiveservices-speech-sdk');

const {
  ActivityHandler,
  CloudAdapter,
  ConfigurationServiceClientCredentialFactory,
  createBotFrameworkAuthenticationFromConfiguration,
} = require('botbuilder');

/** ----------------------------------------------------------------
 * Flowise call helper (unchanged)
 * ---------------------------------------------------------------*/
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

/** ----------------------------------------------------------------
 * Helpers: detect audio in Teams attachments
 * ---------------------------------------------------------------*/
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

function extractAudioInfoFromAttachment(att = {}) {
  const ct = (att.contentType || '').toLowerCase();

  // Teams File Download Info card
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

  // Direct audio content types
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

  // Fallback: filename looks like audio
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

/** ----------------------------------------------------------------
 * Graph token + file download
 * ---------------------------------------------------------------*/
async function getGraphAppToken() {
  const tenant = process.env.MICROSOFT_APP_TENANT_ID;
  const tokenUrl = `https://login.microsoftonline.com/${tenant}/oauth2/v2.0/token`;

  const body = new URLSearchParams({
    grant_type: 'client_credentials',
    client_id: process.env.MICROSOFT_APP_ID,
    client_secret: process.env.MICROSOFT_APP_PASSWORD,
    scope: 'https://graph.microsoft.com/.default',
  });

  const resp = await axios.post(tokenUrl, body.toString(), {
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    timeout: 20_000,
  });

  return resp.data.access_token;
}

async function downloadTeamsProtectedFile(fileUrl, suggestedName = 'audio.wav') {
  const accessToken = await getGraphAppToken();

  const resp = await axios.get(fileUrl, {
    headers: { Authorization: `Bearer ${accessToken}` },
    responseType: 'arraybuffer',
    timeout: 60_000,
    maxRedirects: 5,
    validateStatus: (s) => s >= 200 && s < 400,
  });

  let filename = suggestedName || 'audio.wav';
  if (!path.extname(filename)) filename = `${filename}.wav`;

  const safeName = filename.replace(/[^a-zA-Z0-9._-]/g, '_');
  const stamped = `${Date.now()}_${safeName}`;
  const filePath = path.join('/tmp', stamped);
  fs.writeFileSync(filePath, resp.data);
  return filePath;
}

/** ----------------------------------------------------------------
 * Azure STT (handles m4a/mp3 via ffmpeg → 16k PCM stream)
 * ---------------------------------------------------------------*/
function speechConfig() {
  const cfg = sdk.SpeechConfig.fromSubscription(
    process.env.AZURE_SPEECH_KEY,
    process.env.AZURE_SPEECH_REGION
  );
  cfg.speechRecognitionLanguage = process.env.SPEECH_LANG || 'en-US';
  // Optional: tune end-of-speech; default works fine for short notes
  return cfg;
}

/**
 * Transcribe any audio file by transcoding it with ffmpeg to 16kHz PCM
 * and streaming into Azure Speech SDK.
 */
async function transcribeWithAzure(filePath) {
  return new Promise((resolve, reject) => {
    const pushStream = sdk.AudioInputStream.createPushStream();

    // ffmpeg -> raw s16le 16k mono
    const ff = spawn(ffmpegPath, [
      '-hide_banner',
      '-loglevel', 'error',
      '-i', filePath,
      '-vn',
      '-acodec', 'pcm_s16le',
      '-ac', '1',
      '-ar', '16000',
      '-f', 's16le',
      'pipe:1',
    ]);

    ff.stdout.on('data', (chunk) => pushStream.write(chunk));
    ff.stderr.on('data', (d) => {
      // You can log ffmpeg warnings if needed
    });
    ff.on('close', () => pushStream.close());
    ff.on('error', (e) => reject(new Error('ffmpeg failed: ' + e.message)));

    const audioConfig = sdk.AudioConfig.fromStreamInput(pushStream);
    const recognizer = new sdk.SpeechRecognizer(speechConfig(), audioConfig);

    recognizer.recognizeOnceAsync(
      (result) => {
        recognizer.close();
        resolve(result.text || '');
      },
      (err) => {
        recognizer.close();
        reject(err);
      }
    );
  });
}

/** ----------------------------------------------------------------
 * Bot: now downloads → transcribes → (optionally) asks Flowise
 * ---------------------------------------------------------------*/
class FlowiseBot extends ActivityHandler {
  constructor() {
    super();

    this.onMessage(async (context, next) => {
      const activity = context.activity || {};
      const attachments = activity.attachments || [];

      // Audio path
      const audioItems = [];
      for (const att of attachments) {
        const info = extractAudioInfoFromAttachment(att);
        if (info) audioItems.push(info);
      }

      if (audioItems.length > 0) {
        const audio = audioItems[0];
        console.log('🎤 Audio attachment detected:', audio);

        try {
          await context.sendActivity({ type: 'typing' });

          // 1) Download
          const savedPath = await downloadTeamsProtectedFile(audio.url, audio.name);
          console.log('✅ Saved audio to:', savedPath);

          // 2) Transcribe with Azure
          const transcript = await transcribeWithAzure(savedPath);
          console.log('📝 Transcript:', transcript);

          // 3) Ask Flowise with the transcript
          const flowiseAnswer = transcript
            ? await askFlowise(transcript)
            : '[No speech detected]';

          await context.sendActivity(
            `📝 **Transcript:** ${transcript || '_<empty>_'}\n\n—\n**Answer:** ${flowiseAnswer}`
          );
        } catch (err) {
          console.error('Transcription flow error:', err?.response?.data || err?.message || err);
          await context.sendActivity(
            `⚠️ I detected your audio **${audio.name}**, but I couldn't transcribe it. Please try a different clip or format.`
          );
        }

        await next();
        return;
      }

      // Text path (unchanged)
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

      const reply = { type: 'message', text: answer };
      if (activity.conversation?.conversationType === 'channel') {
        reply.id = activity.id;
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

/** Auth + adapter */
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
