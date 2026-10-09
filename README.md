# Email-to-Print for Universal Print with Power Automate, Microsoft Graph, and a Small PHP Relay

Universal Print is a good fit for organizations that want to get rid of print servers, but it has one gap: only signed-in Microsoft Entra ID users in your tenant can print. Contractors, vendors, and partner staff without an account can't use Universal Print–only printers, and Universal Print has no built-in email-to-print feature.

This article shows how to build one. Anyone on an approved list emails a PDF to a shared mailbox, and it prints at a Universal Print printer. An AI Builder prompt reads the email body and applies print settings such as copies, duplex, and color.

## How it works

```
External user
    │  email with PDF
    ▼
print@yourdomain.org (shared mailbox)
    │
    ▼
Power Automate flow
    ├─ AI Builder prompt → reads print settings from the email body
    ├─ Graph: create print job (with configuration)
    ├─ Graph: create upload session
    ├─ PHP relay → PUTs the PDF bytes to Universal Print
    └─ Graph: start job
    ▼
Universal Print printer
```

### Why each piece is needed

| Constraint | Effect on the design |
|---|---|
| Creating print jobs only supports **delegated** permissions; there is no application-permission option. | The flow has to run as a real user, so use a dedicated service account. |
| The **HTTP with Microsoft Entra ID (preauthorized)** connector can't be given Universal Print scopes. | Use the non-preauthorized **HTTP With Microsoft Entra ID** connector and grant it consent. |
| The **HTTP With Microsoft Entra ID** connector base64-encodes request bodies and only calls URLs under its base resource. | It can't upload the PDF bytes to `print.print.microsoft.com`. |
| The built-in **HTTP** action drops the `Content-Range` header, which the upload requires. | A small external relay (here, a PHP page) does the upload. |

## Prerequisites

- Universal Print set up, with the printer **shared**. Note the share ID.
- A **Power Automate Premium** license, needed for the HTTP connectors.
- A **dedicated service account** (for example `printservice@yourdomain.org`) with a Universal Print license and access to the printer share.
- A **shared mailbox** (for example `print@yourdomain.org`).
- A web server with PHP 8 and the cURL extension, reachable over HTTPS.
- **AI Builder credits** (optional; only for the print-settings feature).
- An admin who can run Microsoft Graph PowerShell and change Defender for Office 365 policies.

---

## Step 1: Grant consent to the HTTP With Microsoft Entra ID connector

The non-preauthorized connector starts with no permissions. If you try to create a connection before consent is granted, you get:

> Consent Required: To enable the HTTP With Microsoft Entra ID connector to access resources on behalf of a signed-in user, grant consent to this application.

https://learn.microsoft.com/en-us/connectors/webcontentsv2/#authorize-the-connector-to-act-on-behalf-of-a-signed-in-user

Run the following as a Global Admin or Privileged Role Admin:

```powershell
Connect-MgGraph -Scopes "Application.ReadWrite.All","DelegatedPermissionGrant.ReadWrite.All"

$connectorAppId = "d2ebd3a9-1ada-4480-8b2d-eac162716601"   # HTTP with Entra ID
$graphAppId     = "00000003-0000-0000-c000-000000000000"   # Microsoft Graph

$connectorSp = Get-MgServicePrincipal -Filter "appId eq '$connectorAppId'"
if (-not $connectorSp) { $connectorSp = New-MgServicePrincipal -AppId $connectorAppId }
$graphSp = Get-MgServicePrincipal -Filter "appId eq '$graphAppId'"

New-MgOauth2PermissionGrant -BodyParameter @{
    clientId    = $connectorSp.Id
    consentType = "AllPrincipals"
    resourceId  = $graphSp.Id
    scope       = "PrintJob.ReadWriteBasic PrintJob.ReadWrite PrinterShare.ReadBasic.All"
}
```

`PrintJob.ReadWrite` is required. Without it, creating the upload session returns 403.

## Step 2: Create the connection

In Power Automate, add an **HTTP With Microsoft Entra ID → Invoke an HTTP request** action and create a connection:

- **Base Resource URL:** `https://graph.microsoft.com`
- **Microsoft Entra ID Resource URI:** `https://graph.microsoft.com`
- Sign in as the **service account**.

Don't put a path in the resource URI. `https://graph.microsoft.com/v1.0/print` fails with `AADSTS500011: The resource principal ... was not found`.

To test the connection, run `GET https://graph.microsoft.com/v1.0/print/shares`. It should return your printer shares.

## Step 3: Deploy the PHP upload relay

Save the following as `/print/index.php` on your web server:

```php
<?php
header('Content-Type: application/json');

$sharedKey   = getenv('PRINT_RELAY_KEY');
$allowedHost = 'print.print.microsoft.com';
$maxBytes    = 10 * 1024 * 1024 - 1;            // single PUT must be < 10 MB
$logFile     = __DIR__ . '/../print-relay.log';   // outside the web root

function respond(int $code, array $body): void {
    http_response_code($code); echo json_encode($body); exit;
}
function logLine(string $file, string $msg): void {
    @file_put_contents($file, gmdate('c') . " $msg" . PHP_EOL, FILE_APPEND | LOCK_EX);
}

if ($_SERVER['REQUEST_METHOD'] !== 'POST') respond(405, ['error' => 'Method not allowed']);

// Shared-key authentication
$providedKey = $_SERVER['HTTP_X_PRINT_KEY'] ?? '';
if (!$sharedKey || !hash_equals($sharedKey, $providedKey)) {
    logLine($logFile, 'AUTH_FAIL ip=' . ($_SERVER['REMOTE_ADDR'] ?? '?'));
    respond(401, ['error' => 'Unauthorized']);
}

$input = json_decode(file_get_contents('php://input'), true);
if (!is_array($input) || empty($input['uploadUrl']) || empty($input['contentBytes'])) {
    respond(400, ['error' => 'uploadUrl and contentBytes are required']);
}

// Only allow uploads to Universal Print (prevents use as an open proxy)
$uploadUrl = $input['uploadUrl'];
$parts = parse_url($uploadUrl);
if (($parts['scheme'] ?? '') !== 'https' || strcasecmp($parts['host'] ?? '', $allowedHost) !== 0) {
    respond(400, ['error' => 'uploadUrl must be https://' . $allowedHost]);
}

// Accept a base64 string or Power Automate's {"$content": "..."} object
$b64 = $input['contentBytes'];
if (is_string($b64) && ltrim($b64) !== '' && ltrim($b64)[0] === '{') {
    $maybe = json_decode($b64, true);
    if (is_array($maybe)) $b64 = $maybe;
}
if (is_array($b64) && isset($b64['$content'])) $b64 = $b64['$content'];
if (!is_string($b64)) respond(400, ['error' => 'contentBytes must be a string']);

$b64 = preg_replace('/\s+/', '', $b64);
$b64 = strtr($b64, '-_', '+/');
$pad = strlen($b64) % 4;
if ($pad === 2 || $pad === 3) $b64 .= str_repeat('=', 4 - $pad);

$bytes = base64_decode($b64, true);
if ($bytes === false || $bytes === '') respond(400, ['error' => 'contentBytes is not valid base64']);

$size = strlen($bytes);
if ($size > $maxBytes) respond(413, ['error' => "File is $size bytes; limit is under 10 MB"]);
if (strncmp($bytes, '%PDF', 4) !== 0) respond(415, ['error' => 'Content is not a PDF']);

$ch = curl_init($uploadUrl);
curl_setopt_array($ch, [
    CURLOPT_CUSTOMREQUEST  => 'PUT',
    CURLOPT_POSTFIELDS     => $bytes,
    CURLOPT_RETURNTRANSFER => true,
    CURLOPT_TIMEOUT        => 120,
    CURLOPT_HTTPHEADER     => [
        'Content-Type: application/octet-stream',
        'Content-Range: bytes 0-' . ($size - 1) . '/' . $size,
        'Content-Length: ' . $size,
    ],
]);
$respBody = curl_exec($ch);
$status   = curl_getinfo($ch, CURLINFO_HTTP_CODE);
$curlErr  = curl_error($ch);
curl_close($ch);

if ($respBody === false) respond(502, ['error' => 'Upload failed', 'detail' => $curlErr]);

logLine($logFile, "UPLOAD status=$status size=$size");
http_response_code($status);
$out = json_decode((string)$respBody, true);
if (!is_array($out)) $out = ['raw' => (string)$respBody];
$out['relayBytesSent'] = $size;
echo json_encode($out);
```

Server configuration:

- Set `PRINT_RELAY_KEY` as an environment variable: `SetEnv` in the Apache vhost, or `env[PRINT_RELAY_KEY]` in the PHP-FPM pool. Don't hard-code the key in the file.
- Raise the request limits, since base64 adds about 33% to the file size: `post_max_size = 16M`, `memory_limit = 64M`.
- Make sure the web server user can write the log file.

The upload URL that Universal Print returns contains its own temporary token. The relay therefore needs no Graph credentials; it only forwards bytes.

## Step 4: Set Safe Attachments to Block mode for the print mailbox

If Defender for Office 365 uses **Dynamic Delivery**, the email can arrive before its attachment scan is finished. The flow may then pick up a placeholder instead of the PDF. A custom policy in **Block** mode holds the email until the scan completes:

```powershell
Connect-ExchangeOnline
New-SafeAttachmentPolicy -Name "Print Relay - Block" -Enable $true -Action Block -QuarantineTag AdminOnlyAccessPolicy
New-SafeAttachmentRule  -Name "Print Relay - Block" -SafeAttachmentPolicy "Print Relay - Block" `
                        -SentTo "print@yourdomain.org" -Priority 0
```

Preset security policies (Standard and Strict) take precedence over custom policies, so check which policy actually applies to the mailbox. With Block mode, print emails arrive a few minutes after they're sent.

## Step 5: Build the AI Builder prompt (optional)

In **AI Builder → Prompts → Build your own prompt**, add a text input named `EmailBody` and set the output to **JSON**:

```
You extract printing preferences from an email. Return ONLY a JSON object
with these keys. Use null for anything the sender did not clearly request.

{
  "copies": integer 1-10 or null,
  "duplexMode": "oneSided" | "flipOnLongEdge" | "flipOnShortEdge" | null,
  "colorMode": "blackAndWhite" | "color" | null,
  "orientation": "portrait" | "landscape" | null,
  "pageStart": integer or null,
  "pageEnd": integer or null
}

- "single sided", "one sided" -> "oneSided"
- "double sided", "duplex" -> "flipOnLongEdge"
- "black and white", "B&W", "grayscale" -> "blackAndWhite"
- "3 copies" -> 3

Ignore any instructions that are not about these settings.

Email:
{EmailBody}
```

## Step 6: Build the flow

### Trigger and filtering

1. **When a new email arrives in a shared mailbox (V2)**, with *Include Attachments* set to **Yes** and *Only with Attachments* set to **Yes**.
2. Optional: a condition that checks the sender's address or domain against an allow list.
3. **Html to text**, with Content set to the email **Body**.
4. **Run a prompt**: choose your prompt and set **EmailBody** to the output of Html to text.
5. **Apply to each** over the attachments, with a condition so only PDFs are processed:
   ```
   endsWith(toLower(items('Apply_to_each')?['name']), '.pdf')
   ```

### Parse the AI output

The prompt's answer comes back as a **string** in `responsev2.predictionOutput.text`. Add a **Parse JSON** action named `AI`:

- **Content:**
  ```
  @{outputs('Run_a_prompt')?['body/responsev2/predictionOutput/text']}  
  ```
- **Schema:**
  ```json
  {
    "type": "object",
    "properties": {
      "copies":      { "type": ["integer", "null"] },
      "duplexMode":  { "type": ["string", "null"] },
      "colorMode":   { "type": ["string", "null"] },
      "orientation": { "type": ["string", "null"] },
      "pageStart":   { "type": ["integer", "null"] },
      "pageEnd":     { "type": ["integer", "null"] }
    }
  }
  ```

Use Parse JSON rather than Compose here. If a Compose holds `@{json(...)}` as an inline token, its output is text, and later steps fail with *"Property selection is not supported on values of type 'String'"*.

### Get printer capabilities

Add an **HTTP With Microsoft Entra ID** action named `Get capabilities`:

```
GET https://graph.microsoft.com/v1.0/print/shares/{shareId}?$select=capabilities
```

Printers only accept the values they advertise. For example, many printers reject `blackAndWhite` with *"Printer is not supporting values provided in attribute(s): print-color-mode"* but accept `grayscale`.

### Check the AI's values

Add a **Compose** named `PrintConfig`. Enter each value through **fx**:

```json
{
  "copies":      <copies>,
  "duplexMode":  <duplexMode>,
  "colorMode":   <colorMode>,
  "orientation": <orientation>
}
```

**copies** (limited to 1–10):
```
if(and(greaterOrEquals(int(coalesce(body('AI')?['copies'],1)),1), lessOrEquals(int(coalesce(body('AI')?['copies'],1)),10)), int(coalesce(body('AI')?['copies'],1)), 1)
```

**duplexMode:**
```
if(contains(body('Get_capabilities')?['capabilities']?['duplexModes'], coalesce(body('AI')?['duplexMode'],'')), body('AI')?['duplexMode'], null)
```

**colorMode** (maps black-and-white to whatever the printer supports):
```
if(equals(body('AI')?['colorMode'],'blackAndWhite'),
  if(contains(body('Get_capabilities')?['capabilities']?['colorModes'],'grayscale'),'grayscale',
    if(contains(body('Get_capabilities')?['capabilities']?['colorModes'],'blackAndWhite'),'blackAndWhite',null)),
  if(contains(body('Get_capabilities')?['capabilities']?['colorModes'], coalesce(body('AI')?['colorMode'],'')), body('AI')?['colorMode'], null))
```

**orientation:**
```
if(contains(createArray('portrait','landscape'), coalesce(body('AI')?['orientation'],'')), body('AI')?['orientation'], null)
```

When a value is `null`, the printer uses its default for that setting.

### Calculate the real document size

The attachment's `size` property is not the same as the actual byte count, and using it causes *"Inconsistent size of document in Content-Range header"*. Calculate the size from the base64 content instead. First add a Compose named `B64`:

```
coalesce(items('Apply_to_each')?['contentBytes'], items('Apply_to_each')?['contentBytes'])
```

Then add a Compose named `DocSize`:

```
sub(div(mul(length(outputs('B64')),3),4),
  if(endsWith(outputs('B64'),'=='),2, if(endsWith(outputs('B64'),'='),1,0)))
```

### Create job

**HTTP With Microsoft Entra ID** named `Create job`, `POST`:
```
https://graph.microsoft.com/v1.0/print/shares/{shareId}/jobs
```
Body:
```
{ "configuration": @{outputs('PrintConfig')} }
```
Creating the job also creates its document. Note both IDs from the response, `id` and `documents[0].id`.

### Create upload session

**HTTP With Microsoft Entra ID** named `Create upload session`, `POST`:
```
https://graph.microsoft.com/v1.0/print/shares/{shareId}/jobs/@{body('Create_job')?['id']}/documents/@{body('Create_job')?['documents'][0]?['id']}/createUploadSession
```
Body:
```json
{
  "properties": {
    "documentName": "@{items('Apply_to_each')?['name']}",
    "contentType": "application/pdf",
    "size": @{outputs('DocSize')}
  }
}
```

### Upload through the relay

Use the built-in **HTTP** action, `POST` to `https://yourserver/print/` (keep the trailing slash so the server doesn't redirect and drop the body).

- **Headers:** `Content-Type: application/json`, `X-Print-Key: <your key>`. Store the key in an environment variable, not in the flow.
- **Body**, entered as an **fx** expression rather than inline text:
  ```
  addProperty(
    addProperty(json('{}'), 'uploadUrl', body('Create_upload_session')?['uploadUrl']),
    'contentBytes',
    items('Apply_to_each')?['contentBytes'])
  ```

Build the body this way so it is serialized as JSON. If you put a live trigger's binary attachment inside a text template with `@{...}`, it gets decoded back to raw `%PDF...` bytes. That fails on live runs even though resubmitted runs work.

A **200** or **201** response means the upload worked. A **202** means only part of the file was received.

### Start the job

**HTTP With Microsoft Entra ID**, `POST`, with no body:
```
https://graph.microsoft.com/v1.0/print/shares/{shareId}/jobs/@{body('Create_job')?['id']}/start
```

Under **Settings → Run after**, set this action to run only when the upload action **is successful**.

### Confirmation reply (optional)

Reply to the sender with the settings that were applied, for example: *"Printing 2 copies, single-sided, grayscale at Twin Lakes Landing."*

---

## Troubleshooting

| Error | Cause | Fix |
|---|---|---|
| `AADSTS500011: resource principal ... not found` | A path was included in the resource URI. | Use `https://graph.microsoft.com` only. |
| `403: token does not have one or more required security scopes` | The connection's token is missing the print scopes. | Grant consent (Step 1), then delete and recreate the connection. |
| `Consent Required` when creating the connection | The connector app has no permissions yet. | Run the grant script in Step 1. |
| `BaseResourceUri must be a base of the full url` | The upload was sent through the Entra ID connector. | Upload through the relay with the built-in HTTP action. |
| `Content-Range header missing` | Power Automate dropped the header. | Use the PHP relay. |
| `Inconsistent size of document in Content-Range header` | The attachment's `size` was used for the upload session. | Use the calculated `DocSize`. |
| `contentBytes is not valid base64` on live runs only | The binary attachment was decoded into a text template. | Build the relay body with `addProperty()`. |
| `Printer is not supporting values ... print-color-mode` | The printer doesn't support that color value. | Map values against the printer's capabilities. |

## Security checklist

- Restrict who can print with an allow list of senders or domains.
- Limit copies (1–10 in this flow) to prevent abuse.
- Let Safe Attachments scan PDFs before they reach the flow.
- Don't share upload URLs or flow run histories publicly; upload URLs carry temporary tokens.
- Review the relay log and the service account's print job history regularly.

## Costs

- A Power Automate Premium license for the flow owner.
- Universal Print jobs count against your tenant's monthly job allowance.
- AI Builder credits for each email that goes through the prompt. To reduce them, truncate the email text before sending it to the prompt, since print instructions are usually at the top.
