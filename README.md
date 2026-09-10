# Signal Desk

Signal Desk brings Explee project `34448` hot leads into a spreadsheet-style workspace. It refreshes every 10 minutes, shows the lead's exact interested reply, email, and phone, and lets a signed-in Microsoft 365 user edit and send a reply through Outlook.

The GitHub Pages site never receives the Explee API key. A small Cloudflare Worker keeps the key private, performs the scheduled sync, and stores leads in D1. Microsoft sign-in uses delegated `Mail.Send`, so a message can only be sent as the person who is currently signed in.

## What is already built

- GitHub Pages deployment workflow
- Excel-style lead table with search, open/sent filter, and detail pane
- Editable recipient, subject, and reply review before sending
- Microsoft 365 sign-in and Outlook `sendMail`
- Draft settings with reusable placeholders
- Explee project-scoped polling every 10 minutes
- Private Explee credential storage and Microsoft token validation
- Responsive phone/tablet layout

## One-time setup

Do these steps after the code is in its GitHub repository.

### 1. Turn on GitHub Pages

1. Open the repository on GitHub.
2. Select **Settings**.
3. In the left sidebar, select **Pages**.
4. Under **Build and deployment**, set **Source** to **GitHub Actions**.
5. Leave this tab open; the page URL will appear here after the first successful deployment.

GitHub's official explanation is [Using custom workflows with GitHub Pages](https://docs.github.com/en/pages/getting-started-with-github-pages/using-custom-workflows-with-github-pages).

### 2. Create the Microsoft 365 connection

1. Open the [Microsoft Entra admin center](https://entra.microsoft.com/).
2. Select **Identity → Applications → App registrations**.
3. Select **New registration**.
4. Name it **Signal Desk**.
5. Choose **Accounts in this organizational directory only**.
6. Leave Redirect URI empty and select **Register**.
7. On the Overview page, copy **Application (client) ID** and **Directory (tenant) ID**.
8. Select **Authentication → Add a platform → Single-page application**.
9. Enter the final GitHub Pages URL, including its trailing slash, then select **Configure**.
10. Select **API permissions → Add a permission → Microsoft Graph → Delegated permissions**.
11. Search for and add **Mail.Send**. `User.Read` is normally already present.

No Microsoft client secret is needed or expected. Microsoft documents the [SPA registration flow](https://learn.microsoft.com/en-us/graph/auth-register-app-v2) and the [Mail.Send delegated permission](https://learn.microsoft.com/en-us/graph/permissions-reference#mail-send).

### 3. Create the secure Cloudflare service

1. Create or sign in to a [Cloudflare account](https://dash.cloudflare.com/).
2. Open **Workers & Pages → D1 SQL database → Create**.
3. Name it **explee-signal-desk** and create it.
4. Copy the database ID.
5. In Cloudflare, open **My Profile → API Tokens → Create Token**.
6. Create a token that can edit Workers Scripts and D1 for this account.
7. Copy the token and your Cloudflare account ID.

### 4. Add safe GitHub settings

In the GitHub repository, open **Settings → Secrets and variables → Actions**.

Under **Variables**, add:

- `MICROSOFT_CLIENT_ID` — the Microsoft Application (client) ID
- `MICROSOFT_TENANT_ID` — the Microsoft Directory (tenant) ID
- `PAGES_ORIGIN` — only the GitHub Pages origin, such as `https://your-name.github.io` (no repository path or trailing slash)
- `API_BASE_URL` — the Worker URL shown after its first deployment

Under **Secrets**, add:

- `EXPLEE_API_KEY` — a newly rotated Explee API key
- `CLOUDFLARE_API_TOKEN`
- `CLOUDFLARE_ACCOUNT_ID`
- `CLOUDFLARE_D1_DATABASE_ID`

The key included in the original request should be rotated before production because it was shared in conversation. Never put it in `.env`, source code, or a GitHub variable.

### 5. Deploy in order

1. Open the repository's **Actions** tab.
2. Select **Deploy secure sync service**.
3. Select **Run workflow**.
4. When it finishes, copy the Worker URL from Cloudflare and save it in the GitHub variable `API_BASE_URL`.
5. Open **Deploy interface to GitHub Pages** in Actions and select **Run workflow**.
6. Open the published page, select **Connect Microsoft 365**, and approve sending mail.
7. Select **Sync now** once. Automatic ten-minute refreshes continue from there.

## Local preview

Copy `.env.example` to `.env.local`, fill only the public Microsoft and Worker values, then run:

```bash
npm install
npm run dev
```

The interface automatically shows clearly fictional preview rows when the live backend is not configured.
