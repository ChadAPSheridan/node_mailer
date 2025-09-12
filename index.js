
require('dotenv').config();
const express = require('express');
const {Client} = require('@microsoft/microsoft-graph-client');
const {ClientSecretCredential} = require('@azure/identity');
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
        const token =
            await credential.getToken('https://graph.microsoft.com/.default');
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
    logToFile(
        `Unauthorized access attempt from IP: ${req.ip}, Key received: ${key}`);
    return res.status(401).json({error: 'Unauthorized'});
  }
  next();
}

app.post('/send-email', authenticateApiKey, async (req, res) => {
  const {to, subject, body, cc, bcc, replyTo, attachments, bodyType} = req.body;
  console.log('Request body:', req.body);
  console.log(
      'To:', to, ', Subject:', subject, ', Body:', body, ', CC:', cc,
      ', BCC:', bcc, ', Reply-To:', replyTo, ', Attachments:', attachments,
      ', Body Type:', bodyType);
  if (!to || !subject || !body) {
    logToFile(`Bad request: Missing required fields from IP: ${req.ip}`);
    return res.status(400).json({error: 'Missing required fields.'});
  }
  try {
    const client = getGraphClient();
    const message = {
      subject,
      body: {contentType: bodyType === 'HTML' ? 'HTML' : 'Text', content: body},
      toRecipients: [{emailAddress: {address: to}}]
    };
    if (cc) {
      const ccList = Array.isArray(cc) ? cc : cc.split(',').map(e => e.trim());
      message.ccRecipients =
          ccList.map(addr => ({emailAddress: {address: addr}}));
    }
    if (bcc) {
      const bccList =
          Array.isArray(bcc) ? bcc : bcc.split(',').map(e => e.trim());
      message.bccRecipients =
          bccList.map(addr => ({emailAddress: {address: addr}}));
    }
    if (replyTo) {
      const replyToList = Array.isArray(replyTo) ?
          replyTo :
          replyTo.split(',').map(e => e.trim());
      message.replyTo =
          replyToList.map(addr => ({emailAddress: {address: addr}}));
    }
    if (attachments && Array.isArray(attachments)) {
      message.attachments = attachments.map(
          att => ({
            '@odata.type': '#microsoft.graph.fileAttachment',
            name: att.name,
            contentBytes: att.contentBytes,
            contentType: att.contentType || 'application/octet-stream'
          }));
    }
    await client.api('/users/' + senderEmail + '/sendMail').post({message});
    res.json({success: true});
  } catch (err) {
    logToFile(`Error sending email: ${err.message} from IP: ${req.ip}`);
    res.status(500).json({error: err.message});
  }
});

app.listen(process.env.PORT || 3000, () => {
  logToFile(`API server started on port ${process.env.PORT || 3000}`);
  console.log(`API server running on port ${process.env.PORT || 3000}`);
});
