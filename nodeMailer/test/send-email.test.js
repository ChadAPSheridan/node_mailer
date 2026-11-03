const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

function loadRoute(graphError) {
  let handler;
  const sent = [];
  const logs = [];
  const app = {
    use() {},
    post(route, authenticate, callback) { handler = callback; },
    listen() {}
  };
  const express = () => app;
  express.json = () => () => {};
  const mocks = {
    dotenv: { config() {} },
    express,
    '@azure/identity': { ClientSecretCredential: class {} },
    '@microsoft/microsoft-graph-client': {
      Client: {
        initWithMiddleware: () => ({
          api: endpoint => ({
            post: async payload => {
              if (graphError) throw graphError;
              sent.push(JSON.parse(JSON.stringify({ endpoint, ...payload })));
            }
          })
        })
      }
    },
    fs: { appendFile: (file, message, callback) => { logs.push(message); callback(); } },
    path
  };
  vm.runInNewContext(fs.readFileSync(path.join(__dirname, '../index.js'), 'utf8'), {
    require: name => mocks[name],
    process: { env: { SENDER_EMAIL: 'sender@example.com' } },
    console: { log() {}, error() {} },
    __dirname: path.join(__dirname, '..')
  });
  return {
    sent,
    logs,
    async send(fields) {
      const res = {
        statusCode: 200,
        status(code) { this.statusCode = code; return this; },
        json(body) { this.body = body; }
      };
      await handler({
        body: { to: 'recipient@example.com', subject: 'Test', body: 'Hello', ...fields },
        ip: '127.0.0.1'
      }, res);
      return res;
    }
  };
}

test('removes the trailing comma from the reported recipient', async () => {
  const route = loadRoute();
  const response = await route.send({ to: 'chad.sheridan@yourvoice.ca,' });
  assert.equal(response.statusCode, 200);
  assert.deepEqual(route.sent[0].message.toRecipients, [
    { emailAddress: { address: 'chad.sheridan@yourvoice.ca' } }
  ]);
});

test('normalizes comma-separated strings and arrays for every recipient field', async () => {
  for (const value of [
    ' chad.sheridan@yourvoice.ca, ,second@example.com, ',
    [' chad.sheridan@yourvoice.ca,', '', ' second@example.com ']
  ]) {
    const route = loadRoute();
    const response = await route.send({ to: value, cc: value, bcc: value, replyTo: value });
    assert.equal(response.statusCode, 200);
    assert.equal(response.body.success, true);
    assert.equal(route.sent[0].endpoint, '/users/sender@example.com/sendMail');
    for (const field of ['toRecipients', 'ccRecipients', 'bccRecipients', 'replyTo']) {
      assert.deepEqual(route.sent[0].message[field], [
        { emailAddress: { address: 'chad.sheridan@yourvoice.ca' } },
        { emailAddress: { address: 'second@example.com' } }
      ]);
    }
  }
});

test('preserves message content and omits empty optional recipient lists', async () => {
  const route = loadRoute();
  const attachments = [{ name: 'test.txt', contentBytes: 'SGVsbG8=' }];
  await route.send({ cc: ' , ', bcc: [], replyTo: null, bodyType: 'HTML', attachments });
  const message = route.sent[0].message;
  assert.deepEqual(message.toRecipients, [{ emailAddress: { address: 'recipient@example.com' } }]);
  assert.deepEqual(message.body, { contentType: 'HTML', content: 'Hello' });
  assert.equal(message.subject, 'Test');
  assert.equal(message.attachments[0].contentBytes, 'SGVsbG8=');
  for (const field of ['ccRecipients', 'bccRecipients', 'replyTo']) {
    assert.equal(field in message, false);
  }
});

test('rejects empty primary recipients and invalid types before sending', async () => {
  for (const fields of [
    { to: ' , ' }, { to: [] }, { to: {} }, { to: [null] },
    { cc: 42 }, { bcc: ['valid@example.com', {}] }, { replyTo: false }
  ]) {
    const route = loadRoute();
    const response = await route.send(fields);
    assert.equal(response.statusCode, 400);
    assert.equal(typeof response.body.error, 'string');
    assert.equal(route.sent.length, 0);
    assert.match(route.logs[0], /Bad request:.*from IP: 127\.0\.0\.1/);
  }
});

test('preserves Graph error responses and logging', async () => {
  const route = loadRoute(new Error('Recipient is not resolved.'));
  const response = await route.send({});
  assert.equal(response.statusCode, 500);
  assert.equal(response.body.error, 'Recipient is not resolved.');
  assert.match(route.logs[0], /Error sending email: Recipient is not resolved/);
});
