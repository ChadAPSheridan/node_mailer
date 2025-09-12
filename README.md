# Node.js Express API for Microsoft 365 Email

This project provides a REST API to send emails using Microsoft Office 365 via the Microsoft Graph API. It uses Express for the server and authenticates using Azure AD credentials (tenant id, client id, client secret).

## Requirements
- API key authentication (`x-api-key` header required)
- Environment variables for Azure and API key

## Features
- POST endpoint to send email
- Uses Microsoft Graph API
- Secure authentication with Azure credentials

## Setup
1. Install dependencies: `npm install`
2. Create a `.env` file with your Azure credentials and API key:
  - TENANT_ID
  - CLIENT_ID
  - CLIENT_SECRET
  - SENDER_EMAIL
  - API_KEY
3. Start the server: `npm start`


## Notes

- Ensure your Azure app has Mail.Send permissions.
- Replace placeholder values in `.env` with your actual credentials.

No build step is required. Run with `npm start` or `node index.js`.

## Persistent Server Setup
### 1. Complete steps 1 and 2 in the [Setup](#setup) section
### 2. Install the [pm2](https://pm2.keymetrics.io/) server
```shell
sudo npm install -g pm2
```
### 3. Start the static server with PM2:
   ```zsh
   pm2 start serve --name nodeMailer -- -s dist -l 3000
   ```
##### *Optionally with a config file:*
1. Create the config file

***pm2.config.json:***
```json
{
  "apps": [
    {
      "name": "nodeMailer",
      "script": "serve",
      "args": "-s dist -l 3000",
      "cwd": "./nodeMailer",
      "env": {
        "PORT": "3000"
      }
    }
  ]
}
```
* <font color=red>*Please note: Adding an env parameter to pm2.config.json will override the corresponding values from your .env file for any variables specified in the env block. PM2 sets these environment variables when starting the process, and they take precedence over those loaded from .env by libraries like dotenv.*

  *Variables not specified in env will still be loaded from .env if your app uses dotenv.*
</font>
2. run the following as root:
```zsh
pm2 start pm2.config.json
```
   - `-s dist` tells `serve` to serve the `dist` folder (the Vite build output).
   - `-l 3000` sets the port (change if needed).

### 4. To make PM2 auto-start on reboot:
   ```zsh
   pm2 save
   pm2 startup
   ```


## Usage
Send a POST request to `/send-email` with JSON body:

```
{
  "to": "destination.contact@example.com",
  "subject": "Test Email",
  "body": "Hello from Office 365 API!",
  "bodyType": "Text", // or "HTML" for HTML emails
  "cc": ["cc1@example.com", "cc2@example.com"], // or "cc1@example.com,cc2@example.com"
  "bcc": ["bcc1@example.com"], // or "bcc1@example.com"
  "replyTo": ["replyto@example.com"], // or "replyto@example.com"
  "attachments": [
    {
      "name": "test.txt",
      "contentBytes": "SGVsbG8gd29ybGQh",  // base64 encoded content
      "contentType": "text/plain"
    }
  ]
}
```

**Note:**
- `cc`, `bcc`, and `replyTo` can be either arrays or comma-separated strings. Both formats are accepted:
  - Array: `["cc1@example.com", "cc2@example.com"]`
  - String: "cc1@example.com,cc2@example.com"
- The API requires the `x-api-key` header in all requests.
- The examples below assume the server is running on port 3000,

### Sending HTML Email

To send an HTML email, set `bodyType` to `HTML` and provide your HTML content in the `body` field:

```
{
  "to": "destination.contact@example.com",
  "subject": "HTML Test",
  "body": "<h1>Hello!</h1><p>This is an <b>HTML</b> email.</p>",
  "bodyType": "HTML"
}
```


## Full Code Samples for Sending Email with Attachment

### PHP
```php
$file = 'test.txt';
$content = base64_encode(file_get_contents($file));
$data = [
    // Build the attachments array for the API payload
    $attachments = [];
    foreach ($input['attachments'] as $filePath) {
      $attachments[] = [
        "name" => basename($filePath),
        "contentBytes" => base64_encode(file_get_contents($filePath)),
        "contentType" => mime_content_type($filePath)
      ];
    }

    // Example payload construction
    $data = [
      "to" => $input['to'],
      "subject" => $input['subject'],
      "body" => $input['body'],
      "cc" => $input['cc'],
      "bcc" => $input['bcc'],
      "attachments" => $attachments
    ];
];
$ch = curl_init('http://localhost:3000/send-email');
curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);
curl_setopt($ch, CURLOPT_POST, true);
curl_setopt($ch, CURLOPT_HTTPHEADER, [
    'Content-Type: application/json',
    'x-api-key: your-api-key'
]);
curl_setopt($ch, CURLOPT_POSTFIELDS, json_encode($data));
$response = curl_exec($ch);
curl_close($ch);
echo $response;
```

### Bash Script
```bash
file="test.txt"
content=$(base64 -w 0 "$file")
curl -X POST http://localhost:3000/send-email \
  -H "Content-Type: application/json" \
  -H "x-api-key: your-api-key" \
  -d '{
    "to": "destination.contact@example.com",
    "subject": "Test Email",
    "body": "Hello from Office 365 API!",
    "attachments": [{
      "name": "test.txt",
      "contentBytes": "'$content'",
      "contentType": "text/plain"
    }]
  }'
```

### Java
```java
import java.io.*;
import java.net.*;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.Base64;

public class SendEmail {
    public static void main(String[] args) throws Exception {
        String filePath = "test.txt";
        byte[] fileContent = Files.readAllBytes(Paths.get(filePath));
        String base64Content = Base64.getEncoder().encodeToString(fileContent);
        String json = "{"
            + "\"to\": \"destination.contact@example.com\"," 
            + "\"subject\": \"Test Email\"," 
            + "\"body\": \"Hello from Office 365 API!\"," 
            + "\"attachments\": [{"
            + "\"name\": \"test.txt\"," 
            + "\"contentBytes\": \"" + base64Content + "\"," 
            + "\"contentType\": \"text/plain\"}]}";
        URL url = new URL("http://localhost:3000/send-email");
        HttpURLConnection con = (HttpURLConnection) url.openConnection();
        con.setRequestMethod("POST");
        con.setRequestProperty("Content-Type", "application/json");
        con.setRequestProperty("x-api-key", "your-api-key");
        con.setDoOutput(true);
        try(OutputStream os = con.getOutputStream()) {
            byte[] input = json.getBytes("utf-8");
            os.write(input, 0, input.length);
        }
        try(BufferedReader br = new BufferedReader(new InputStreamReader(con.getInputStream(), "utf-8"))) {
            StringBuilder response = new StringBuilder();
            String responseLine;
            while ((responseLine = br.readLine()) != null) {
                response.append(responseLine.trim());
            }
            System.out.println(response.toString());
        }
    }
}
```

### TypeScript
```typescript
import { readFileSync } from 'fs';
import fetch from 'node-fetch';

const file = 'test.txt';
const content = readFileSync(file).toString('base64');

const body = {
  to: 'destination.contact@example.com',
  subject: 'Test Email',
  body: 'Hello from Office 365 API!',
  attachments: [{
    name: 'test.txt',
    contentBytes: content,
    contentType: 'text/plain'
  }]
};

fetch('http://localhost:3000/send-email', {
  method: 'POST',
  headers: {
    'Content-Type': 'application/json',
    'x-api-key': 'your-api-key'
  },
  body: JSON.stringify(body)
})
  .then(res => res.json())
  .then(console.log)
  .catch(console.error);
```
