# Microsoft Entra SSO and HTTPS Setup

The Entra sign-in callback must use a public HTTPS origin. If you already have a reverse proxy or tunnel that terminates TLS, you can keep the app itself on HTTP and point `PUBLIC_BASE_URL` at that public HTTPS origin.

For direct local development, use the bundled HTTPS listener in `example/server.ts`.

## 1. Generate a local certificate

From the `example/` folder:

```bash
npm run generate:https-cert
```

This creates a local `.pfx` bundle under `example/certs/` and prints the matching environment values.

## 2. Enable the HTTPS listener

Set these values in `example/.env`:

```env
HTTPS_ENABLED=true
HTTPS_PFX_FILE=certs/localhost.pfx
PUBLIC_BASE_URL=https://localhost:3030
```

If you already have PEM files, you can use:

```env
HTTPS_ENABLED=true
HTTPS_CERT_FILE=certs/localhost.crt
HTTPS_KEY_FILE=certs/localhost.key
PUBLIC_BASE_URL=https://localhost:3030
```

Optional:

- `HTTPS_PORT` changes the HTTPS listener port.
- `HTTPS_PASSPHRASE` protects the PFX or private key if one is required.
- `HTTPS_CA_FILE` lets you provide a custom CA chain.

Start the example app from `example/` with `npm start`.

## 3. Register the Microsoft Entra app

In the Entra app registration:

- Add a **Web** redirect URI that matches the public callback exactly:

```text
https://your-host/api/auth/entra/callback
```

- Make sure the host in the redirect URI matches `PUBLIC_BASE_URL`.
- Use the OIDC scopes `openid profile email`.
- No Microsoft Graph permissions are required for sign-in.
- If you want UPN fallback mapping, add `upn` as an optional claim.

## 4. Configure the app

In the viewer, open **Settings** and select **Microsoft Entra SSO**. Enter:

- Tenant ID
- Client ID
- Client secret

The secret stays server-side and is never returned to the browser.

Entra identities are mapped to an existing local user by normalized email-like values:

- `email`
- `preferred_username`
- `upn`

The matched local account is the one already stored in the viewer.

## 5. Troubleshooting

- `AADSTS900971` usually means the redirect URI does not exactly match the configured Entra reply URL, or `PUBLIC_BASE_URL` is missing/wrong.
- If Microsoft sign-in works through a reverse proxy but not on localhost, enable the HTTPS listener or expose the app through a public HTTPS origin.
- Self-signed local certificates will trigger browser trust warnings unless you trust the generated cert.
