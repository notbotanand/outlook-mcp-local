# Microsoft App Registration Setup

`outlook-mcp-local` requires a Microsoft application/client ID that you own.

The project does not ship a shared Microsoft OAuth application.

## Create the App

1. Open Microsoft app registrations:

```text
https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade
```

2. Choose **New registration**.

3. Set the name:

```text
outlook-mcp-local
```

4. For supported account types, choose:

```text
Personal Microsoft accounts only
```

5. Configure the redirect URI shown by setup:

```text
http://127.0.0.1:53682/callback
```

Use a public client / mobile and desktop style platform. Do not create or use a client secret.

6. Copy the **Application (client) ID** and paste it into:

```bash
npm run setup
```

The setup wizard will not print the configured client ID after it is saved.

## API Permissions

Use **Delegated permissions only**.

Default read mode needs:

```text
Mail.Read
Calendars.Read
```

Calendar-write mode needs:

```text
Mail.Read
Calendars.ReadWrite
```

Do not add Application permissions.

Do not add email write permissions such as:

```text
Mail.Send
Mail.ReadWrite
MailboxSettings.ReadWrite
```

## Public Client Notes

OAuth uses authorization code with PKCE. The app is a public desktop/local client and does not need a client secret.

If Microsoft rejects auth with an error saying the client is not supported for this feature, confirm the app has a localhost redirect URI matching the setup wizard output.
