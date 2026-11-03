
require('dotenv').config();
const express = require('express');
const { Client } = require('@microsoft/microsoft-graph-client');
const { ClientSecretCredential } = require('@azure/identity');
const fs = require('fs');
const path = require('path');

const app = express();
app.use(express.json());

const tenantId = process.env.TENANT_ID;
const clientId = process.env.CLIENT_ID;
const clientSecret = process.env.CLIENT_SECRET;
const senderEmail = process.env.SENDER_EMAIL;
const apiKey = process.env.API_KEY;

const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);

function getGraphClient() {
  return Client.initWithMiddleware({
    authProvider: {
      getAccessToken: async () => {
        const token = await credential.getToken('https://graph.microsoft.com/.default');
        return token.token;
      }
    }
  });
}

// Logging utility
function logToFile(message) {
  const logPath = path.join(__dirname, 'logs', 'app.log');
  const timestamp = new Date().toISOString();
  fs.appendFile(logPath, `[${timestamp}] ${message}\n`, err => {
    if (err) console.error('Failed to write to log file:', err);
  });
}

// API key authentication middleware
function authenticateApiKey(req, res, next) {
  const key = req.headers['x-api-key'];
  if (!key || key !== apiKey) {
    logToFile(`Unauthorized access attempt from IP: ${req.ip}, Key received: ${key}`);
    return res.status(401).json({ error: 'Unauthorized' });
  }
  next();
}

function normalizeRecipients(value, field) {
  if (value === undefined || value === null) return [];
  const values = Array.isArray(value) ? value : [value];
  if (values.some(address => typeof address !== 'string')) {
    throw new TypeError(`${field} must be a string or an array of strings.`);
  }
  return values
    .flatMap(address => address.split(','))
    .map(address => address.trim())
    .filter(Boolean)
    .map(address => ({ emailAddress: { address } }));
}

app.post('/send-email', authenticateApiKey, async (req, res) => {
  const { to, subject, body, cc, bcc, replyTo, attachments, bodyType } = req.body;
  console.log('Request body:', req.body);
  console.log('To:', to, ", Subject:", subject, ", Body:", body, ", CC:", cc, ", BCC:", bcc, ", Reply-To:", replyTo, ", Attachments:", attachments, ", Body Type:", bodyType);
  if (!to || !subject || !body) {
    logToFile(`Bad request: Missing required fields from IP: ${req.ip}`);
    return res.status(400).json({ error: 'Missing required fields.' });
  }
  let toRecipients, ccRecipients, bccRecipients, replyToRecipients;
  try {
    toRecipients = normalizeRecipients(to, 'to');
    ccRecipients = normalizeRecipients(cc, 'cc');
    bccRecipients = normalizeRecipients(bcc, 'bcc');
    replyToRecipients = normalizeRecipients(replyTo, 'replyTo');
    if (!toRecipients.length) {
      throw new TypeError('to must contain at least one recipient.');
    }
  } catch (err) {
    logToFile(`Bad request: ${err.message} from IP: ${req.ip}`);
    return res.status(400).json({ error: err.message });
  }
  try {
    const client = getGraphClient();
    const message = {
      subject,
      body: {
        contentType: bodyType === 'HTML' ? 'HTML' : 'Text',
        content: body
      },
      toRecipients
    };
    if (ccRecipients.length) {
      message.ccRecipients = ccRecipients;
    }
    if (bccRecipients.length) {
      message.bccRecipients = bccRecipients;
    }
    if (replyToRecipients.length) {
      message.replyTo = replyToRecipients;
    }
    if (attachments && Array.isArray(attachments)) {
      message.attachments = attachments.map(att => ({
        '@odata.type': '#microsoft.graph.fileAttachment',
        name: att.name,
        contentBytes: att.contentBytes,
        contentType: att.contentType || 'application/octet-stream'
      }));
    }
    await client.api('/users/' + senderEmail + '/sendMail').post({ message });
    res.json({ success: true });
  } catch (err) {
    logToFile(`Error sending email: ${err.message} from IP: ${req.ip}`);
    res.status(500).json({ error: err.message });
  }
});
let port = process.env.PORT || 3000;
app.listen(port, () => {
  logToFile('API server started on port '+port);
  console.log('API server running on port '+port);
});
