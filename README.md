# Ask Claude for Excel

Adds a real Excel function, `=CLAUDE.ASK("question")`, that calls Claude and drops the
answer straight into the cell — plus a small settings pane (opened from a "Claude
Settings" button on the Home tab) where you paste your API key and pick a model.

There's no floating bar or custom chrome. You type into the formula bar like any other
formula, which is the point: it disappears into Excel until you use it.

## What's in here

| File | Purpose |
|---|---|
| `manifest.xml` | Tells Excel this add-in exists, where its files live, and registers the `CLAUDE` function namespace. |
| `functions.js` / `functions.json` / `functions.html` | The `ASK` function itself (hidden background runtime). |
| `taskpane.html` / `taskpane.js` | The settings pane: API key, model choice, and a "try it here" box. |
| `commands.html` / `commands.js` | Small stub some Excel versions expect to exist; not doing anything here. |
| `assets/` | Icons. |

## 1. Get an Anthropic API key

Go to [console.anthropic.com/settings/keys](https://console.anthropic.com/settings/keys),
sign in, and create a key. You'll paste it into the add-in's settings pane later — it's
stored inside the workbook (via Excel's own settings API), not in this project, and every
request goes straight from your browser/Excel to `api.anthropic.com` — nothing passes
through a third-party server. Keep in mind this also means anyone who opens that workbook
file can find the key if they dig into it, so don't send workbooks with a saved key to
other people.

Note: with your key stored in Excel's client-side add-in, you pay Anthropic directly for
usage (per the API's per-token pricing) — there's no subscription bundling here.

## 2. Host the files somewhere with HTTPS

Excel loads add-in files over the network even when "sideloaded" for personal use, so
they need to live at a real HTTPS URL — `file://` paths don't work. The free, simplest
option is **GitHub Pages**:

1. Create a new GitHub repo (public or private, either is fine), e.g. `claude-excel-addin`.
2. Push every file in this folder to that repo (root of the `main` branch — don't nest
   it in a subfolder).
3. In the repo's Settings → Pages, set the source to "Deploy from branch", branch `main`,
   folder `/root`. Save.
4. Wait a minute or two, then note the URL GitHub gives you, something like
   `https://YOUR-USERNAME.github.io/claude-excel-addin/`.

Any other static HTTPS host works too (Cloudflare Pages, Netlify, Vercel, Azure Static
Web Apps) if you'd rather use one of those.

### Point the manifest at your real URL

`manifest.xml` currently has the placeholder domain
`https://YOUR-USERNAME.github.io/claude-excel-addin` in eight places. Replace all of them
with your actual Pages URL. Easiest via a terminal, from inside this folder:

```bash
sed -i '' 's#https://YOUR-USERNAME.github.io/claude-excel-addin#https://chuck.github.io/claude-excel-addin#g' manifest.xml
```

(that example replaces the placeholder with `https://chuck.github.io/claude-excel-addin`
— swap in your actual GitHub username and repo name; on Linux drop the `''` right after
`-i`). Or just open `manifest.xml` in a text editor and find-and-replace the placeholder
string once. Push the updated file.

## 3. Sideload it into Excel

**Excel on Windows or Mac (desktop):**
Insert tab → Add-ins → My Add-ins → the "..." or gear icon → Upload My Add-in → pick your
local copy of `manifest.xml`.

**Excel on the web:**
Insert tab → Add-ins → Upload My Add-in → Browse → select `manifest.xml`.

Once loaded, click **Claude Settings** on the Home tab, paste your API key, pick a model,
hit Save. Then in any cell:

```
=CLAUDE.ASK("What's a catchy one-line tagline for a coffee shop?")
=CLAUDE.ASK("Summarize this in 5 words", A1:A20)
```

The second form passes a cell or range in as context — Claude sees the raw values from
those cells along with your question.

## Notes and limits

- **First load is slow-ish.** Each call is a live API request (typically 1-5 seconds).
  Don't drag `=CLAUDE.ASK(...)` down thousands of rows at once — you'll fire that many
  API calls and it'll be slow and cost real money.
- **Recalculation:** Excel treats this as a normal (non-volatile) function, so it only
  re-runs when its inputs change or you force a recalc, not on every keystroke elsewhere
  in the sheet.
- **Errors show as Excel errors** (e.g. "Invalid API key") right in the cell if something's
  wrong — hover the cell for the detail.
- **Changing the model list:** edit the `<select>` options in `taskpane.html` if newer
  Claude models come out; check current model IDs at
  [platform.claude.com/docs/en/about-claude/models/overview](https://platform.claude.com/docs/en/about-claude/models/overview).
