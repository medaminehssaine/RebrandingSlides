# Google Drive Setup (5 minutes)

## Step 1 — Create a Google Cloud project

1. Go to https://console.cloud.google.com
2. Click **New Project** → name it anything → **Create**
3. Make sure it's selected in the top dropdown

## Step 2 — Enable the Drive API

1. Go to **APIs & Services → Library**
2. Search "Google Drive API" → click it → **Enable**

## Step 3 — Create OAuth credentials

1. Go to **APIs & Services → Credentials**
2. Click **+ Create Credentials → OAuth client ID**
3. If prompted, configure the consent screen first:
   - User type: **External** → fill in app name (anything) → Save
4. Application type: **Desktop app** → name it → **Create**
5. Click **Download JSON** → save as `credentials.json` in this folder

## Step 4 — Authenticate

1. Start the app: `npm start`
2. Look for a URL printed in the terminal like:
   `⚠️  Drive not authenticated. Visit: https://accounts.google.com/o/oauth2/auth?...`
3. Open that URL in your browser → sign in → Allow
4. You'll be redirected to `localhost:3000/oauth2callback`
5. You'll see ✅ Google Drive connected

That's it! Token is saved in `token.json` — you won't need to do this again.

---

**Files created in this folder:**
- `credentials.json` — your OAuth client (keep private, don't commit)
- `token.json` — your access token (auto-created, keep private)
