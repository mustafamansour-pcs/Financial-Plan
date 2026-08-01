# Household Ledger — Setup Guide

Three things to do, in order: prepare the Google Sheet, get Google Cloud
credentials, then publish the app on GitHub Pages. About 20–30 minutes
total, all free.

---

## 1. Prepare your Google Sheet

1. Upload your `Home.xlsx` to Google Sheets (File → Open → Upload, or
   drag it into Google Drive and open with Google Sheets). Keep the tab
   names exactly as they are: `Cash Flow`, `Actuals`, `Budget`,
   `Budget Breakdown`.
2. Add one new tab called **`Users`**. In row 1 put headers, then one row
   per family member:

   | Name  | PIN  | Google Email          | Allowed Categories |
   |-------|------|------------------------|---------------------|
   | You   | 4821 | you@gmail.com          | ALL                 |
   | Mona  | 1193 | mona@gmail.com         | Home Expenses       |
   | Sam   | 7304 | sam@gmail.com          | Car Expenses,Outings|

   - **PIN**: any 4 digits, unique per person.
   - **Allowed Categories**: comma-separated, must match category names
     in your sheet exactly. Use `ALL` for full admin access (only you
     should have this).
3. Copy your **Sheet ID** from the URL:
   `https://docs.google.com/spreadsheets/d/`**`THIS_LONG_ID`**`/edit`
4. Click **Share** → add each family member's Google email as **Editor**.
   This is what lets Google Sign-In authorize their write requests.

---

## 2. Google Cloud credentials (free)

1. Go to [console.cloud.google.com](https://console.cloud.google.com) and
   create a new project (top-left project dropdown → New Project). Name
   it anything, e.g. "Household Ledger".
2. **Enable the Sheets API**: in the search bar, type "Google Sheets API"
   → open it → click **Enable**.
3. **Create the API key** (for reading the Top Sheet):
   - APIs & Services → Credentials → **Create Credentials** → **API key**
   - Click the new key to edit it → under "API restrictions" choose
     **Restrict key** → select **Google Sheets API** only → Save.
   - Copy this key into `config.js` as `API_KEY`.
4. **Create the OAuth Client ID** (for signed-in writes):
   - Credentials → **Create Credentials** → **OAuth client ID**
   - If prompted, configure the consent screen first: choose **External**,
     fill in an app name and your email, save through the defaults. You
     don't need to submit for verification for personal/family use —
     add each family member's email under "Test users" on that screen.
   - Application type: **Web application**
   - Under "Authorized JavaScript origins" add the URL your app will
     live at, e.g. `https://yourusername.github.io` (add it once you
     know your GitHub Pages URL from step 3 below — you can come back
     and edit this).
   - Copy the Client ID into `config.js` as `OAUTH_CLIENT_ID`.
5. Fill in `config.js` fully:
   ```js
   SHEET_ID: "your-sheet-id",
   API_KEY: "your-api-key",
   OAUTH_CLIENT_ID: "your-client-id.apps.googleusercontent.com",
   ```

---

## 3. Publish on GitHub Pages (free)

1. Create a new GitHub repository (e.g. `household-ledger`), public or
   private — Pages works with either on a personal GitHub account,
   though private repos need GitHub Pro for Pages. Public is fine for
   this since there's no secret in the code (the API key is read-only
   and restricted to Sheets API; write access always requires each
   person's own Google sign-in).
2. Upload all the files from this project (`index.html`, `style.css`,
   `app.js`, `data.js`, `config.js`, `manifest.json`) to the repo —
   either drag-and-drop in the GitHub web UI, or:
   ```bash
   git init
   git add .
   git commit -m "Household ledger app"
   git branch -M main
   git remote add origin https://github.com/YOURUSERNAME/household-ledger.git
   git push -u origin main
   ```
3. In the repo: **Settings → Pages** → under "Build and deployment",
   source = **Deploy from a branch**, branch = **main**, folder = **/(root)**.
   Save.
4. After a minute, your app is live at:
   `https://YOURUSERNAME.github.io/household-ledger/`
5. Go back to Google Cloud Console → your OAuth Client ID → add that
   exact URL under "Authorized JavaScript origins" (no trailing slash).

---

## 4. Try it

1. Open the URL on your iPhone in Safari.
2. Enter a PIN from your Users tab.
3. On **Add Entry**, tap **Sign in with Google** the first time — use
   the same email you shared Editor access with.
4. Add a test entry, then check it landed in the `Actuals` tab.
5. Tap **Add to Home Screen** in Safari's share sheet so it behaves like
   a real app icon going forward.

---

## Notes and known limits

- **Refresh model**: the Top Sheet re-fetches on open, after you submit
  an entry, and via the refresh button — not instant push. This matches
  what we agreed (Option A).
- **Admin permission edits**: tapping category chips in the Admin tab
  currently only *stages* a visual change and reminds you to edit the
  `Users` tab directly in Google Sheets — wiring that button to actually
  write back needs the same OAuth write flow as Add Entry, applied to
  the Users tab. Say the word and I'll add it.
- **Cash Flow parsing**: the app reads category rows by matching the
  "Budget" / "Actual Cost" labels in column B, the same way your sheet
  is laid out today. If you restructure that sheet later, the parsing
  logic in `app.js` (`loadTopSheet`) will need a matching update.
- **Item list**: the Add Entry item field is free text rather than a
  dropdown, since your sheet's Item list is partial and subcategory-
  dependent. Happy to wire up the full dependent dropdown if useful.
