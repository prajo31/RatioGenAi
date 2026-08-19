# Real-Company Ratio Comparison Tool (with student login)

This is your existing ratio-analysis Streamlit app (`AI Gemini.py`) with a
login screen added in front of it: students must sign in with a
username/password before they can use the tool, and each generated
CSV/Excel/PDF report is stamped "Prepared by: <student name>".

## Files

- `app.py` — the app itself (login gate + your original ratio comparison tool, unchanged below the login).
- `build_roster.py` — turns a plain-text roster CSV into `config.yaml` (hashed passwords).
- `roster.csv` — **sample** roster with 2 demo logins. Replace with your real class list.
- `config.yaml` — generated login file (hashed passwords). Already built from the sample roster so you can try it immediately.
- `requirements.txt` — everything needed to run the app.

## Try it locally right now

```bash
pip install -r requirements.txt
streamlit run app.py
```

Log in with one of the sample accounts from `roster.csv`:

| username | password |
|---|---|
| demo_student | ChangeMe123! |
| jdoe | ChangeMe123! |

## Setting up your real class roster

1. Open `roster.csv` and replace the sample rows with your actual students —
   one row per student: `username,name,email,password`.
   - `username` is what they type to log in (e.g. a student ID or their
     first initial + last name). Must be unique per student.
   - `password` is a starting password you assign. Students can't change it
     from inside the app — to reset someone's password, edit their row and
     rerun the build script.
2. Regenerate the login file:
   ```bash
   python build_roster.py
   ```
   This overwrites `config.yaml` with freshly hashed passwords for everyone
   in `roster.csv`.
3. **Never share or commit `roster.csv` or `config.yaml` publicly** —
   `roster.csv` has plain-text passwords, and `config.yaml` has the hashed
   versions plus a cookie secret. Keep both private (e.g. add them to
   `.gitignore` if this becomes a git repo).

## Deploying to Streamlit Community Cloud

Don't upload `config.yaml` to a public repo. Instead:

1. Push `app.py`, `build_roster.py`, and `requirements.txt` to a repo
   (leave `roster.csv`/`config.yaml` out, or keep the repo private).
2. Deploy the app on share.streamlit.io.
3. Run `python build_roster.py` locally, and copy the TOML block it prints
   at the end into the app's **Settings → Secrets** box on Streamlit Cloud.
   That gives the deployed app the same login roster without it ever
   touching your public repo.
4. Redeploy/reboot the app after saving secrets.

## AI interpretation: bring-your-own-chat, no API key

This app does **not** call any AI API and never asks for a key — nothing
about a student's AI account touches this app or your server. Consumer
chat products (ChatGPT, Gemini, Claude.ai, Perplexity) don't allow being
embedded in another site, and there's no supported way for a third-party
app to drive a student's already-logged-in chat session for them, so
instead the Summary Dashboard tab gives students a three-step manual
bridge:

1. **Open your own AI chat** — buttons link out to chatgpt.com,
   gemini.google.com, claude.ai, and perplexity.ai, each opening in a new
   browser tab using whatever account the student is already logged into.
2. **Copy the prompt** — the app builds a prompt from the already-computed
   ratio table (never raw financials) and shows it in a copyable code box.
3. **Paste the answer back** — the student pastes the AI's response into a
   text box in the app, which renders it alongside the ratios, includes the
   original prompt for disclosure, and requires the student to write a
   short verification note before treating it as done.

This keeps everything on the student's own AI account and quota — no
shared instructor key, no API cost tracking, no dependency on any one AI
provider.

## New: AI Ratio Calculation Challenge tab (two rounds)

This is a second, separate AI exercise — additive, doesn't touch anything
else in the app. The Summary Dashboard's "AI Interpretation" feature (above)
hands the AI ratios the app already calculated and asks it to comment on
them. The new **🧪 AI Calc Challenge** tab instead tests whether the AI can
calculate — and even research — the ratios itself, across two rounds per
company.

**Round 1 — AI calculates from data you give it.** The AI gets the exact
same raw balance sheet/income statement figures the app itself used
(post-edit, same numbers behind that company's ratios) and has to calculate
every ratio itself, showing its formula and work. Because the input data is
identical to the app's, any mismatch can only mean one thing: the AI's
arithmetic or formula was wrong. The student copies the prompt, pastes the
AI's response back for disclosure, transcribes its "Final Answers" table,
and the app automatically compares those numbers against its own
yfinance-based values (5% tolerance) with a Match/Differ verdict.

**Round 2 — open research, four-way comparison.** Nobody gets handed any
figures.

- **Step 0 — predict first.** Before looking anything up, the student fills
  in their own gut-check guess for every ratio in an editable table, then
  clicks "I've made my predictions — unlock the rest of Round 2." This is a
  hard gate, not an honor-system note: the manual 10-K lookup, the AI
  prompt, and the comparison stay hidden until the student commits to a
  prediction, so nobody can quietly fill in a "guess" that actually just
  matches what they saw below. This is meant to build real engagement
  instead of passive copy-pasting.
- **Step 1 — manual research.** The student looks up the company's actual
  10-K/annual report (e.g. via SEC EDGAR or investor relations) and enters
  what they find into an editable table — real extra work, not just
  re-reading the app's own data.
- **Step 2 — AI research.** Separately, the AI is given a prompt with *no
  data at all* and has to research the same company on its own (training
  knowledge or its own browsing/search), explicitly stating which fiscal
  period and source it used before showing its calculations — that
  requirement is what makes a mismatch diagnosable afterward (data problem
  vs. math problem vs. both) rather than just "wrong."

The app then shows all four answers side by side — Your Prediction, App
(yfinance), Manual (10-K), AI (self-sourced) — flags each relevant pair for
a match/difference (You vs. App, App vs. Manual, App vs. AI), and asks the
student to reflect on where and why they diverge, and which source they'd
actually trust for a real assignment. An optional "📊 Show per-ratio
comparison charts" toggle renders a small grouped bar chart (You / App /
Manual / AI) for every ratio, for a quicker visual read than the table
alone.

It's a verification exercise, not a trust exercise: AI models are
generally solid at explaining ratios but do sometimes get the data or the
math wrong, especially on trickier formulas (interest coverage, DSO,
book-to-market) or when sourcing their own figures — this tab is built to
catch that, not to assume it away.

**AI use disclosure.** Under both Round 1's and Round 2's pasted-response
box, the app asks "Which AI did you use?" (ChatGPT / Gemini / Claude /
Perplexity / Other) and automatically stamps a "Recorded: `<date time>`"
timestamp the first time a response is pasted (and updates it if the
student edits the pasted text later). This travels into the Round 2 PDF
export automatically, so the instructor sees which tool was used and when,
without relying on the student to remember to mention it.

**Submitting Round 2 for grading.** Each company's Round 2 section ends
with two download buttons: a CSV of just the comparison table (quick
numbers-only record), and a PDF that bundles everything an instructor
would need to grade it — the four-way table (including the student's own
prediction), the student's 10-K source citation, the exact prompt sent to
the AI, the AI's full pasted response with its disclosed tool/timestamp,
and the student's written reflection, all in one landscape-orientation
document sized to fit the extra columns.

## Persistent progress (optional — Google Sheets)

By default, everything a student types in the **AI Calc Challenge** tab only
lives in that browser tab for the current session — closing the tab, logging
out, or (on Streamlit Community Cloud) the app simply going to sleep and
waking back up can lose unsaved work if they haven't downloaded their
CSV/PDF yet. If you want students to be able to log out and pick up later
where they left off, you can turn on autosave to a Google Sheet — free, no
server of your own to run, and it reuses the exact same `st.secrets`
mechanism you already use for login credentials.

**Setup (one-time, ~10 minutes):**

1. In the [Google Cloud Console](https://console.cloud.google.com/), create
   a project (or reuse one), then enable the **Google Sheets API** for it.
2. Create a **Service Account** (APIs & Services → Credentials → Create
   Credentials → Service Account), then create a JSON key for it and
   download it. Note the `client_email` address inside that JSON file —
   it looks like `something@your-project.iam.gserviceaccount.com`.
3. Create a new, blank Google Sheet (any name), open it, click **Share**,
   and share it with that service account's email as **Editor**. Copy the
   Sheet's ID out of its URL — the long string between `/d/` and `/edit`.
4. Add both to your secrets — either `.streamlit/secrets.toml` locally
   (already git-ignored) or Streamlit Community Cloud's **Settings →
   Secrets**, alongside your existing `[credentials]`/`[cookie]` blocks:
   ```toml
   progress_sheet_id = "paste-the-sheet-id-here"

   [gcp_service_account]
   type = "service_account"
   project_id = "your-project-id"
   private_key_id = "..."
   private_key = "-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----\n"
   client_email = "something@your-project.iam.gserviceaccount.com"
   client_id = "..."
   token_uri = "https://oauth2.googleapis.com/token"
   ```
   Every field above (except `progress_sheet_id`, which you copy from the
   Sheet's URL) can be copy-pasted straight out of the downloaded JSON key
   file — same keys, just reformatted as TOML.
5. Redeploy/reboot the app. The app creates a `progress` worksheet inside
   that Sheet automatically the first time it's needed.

**What gets saved.** For each student and company, one row keyed by
username/ticker/year holds their Round 1 and Round 2 answers (pasted AI
responses, transcribed values, predictions, 10-K source citation, AI
tool/timestamp disclosure, reflections) as they type — no separate "Save"
button required. A small "💾 Progress autosaved at `<time>`" line appears
under each company once it's working. If Sheets isn't configured, or a save
attempt fails, the app says so plainly instead of failing silently, and
everything still works normally for the current browser session — students
should just remember to download their CSV/PDF before closing the tab in
that case.

**Note for instructors:** since this Sheet accumulates students' pasted AI
responses and reflections, treat it like any other student-work record
(don't make it public, and follow your institution's data retention norms).

## What changed from the original app

- Added a login screen (`streamlit-authenticator`) gating the whole tool —
  nothing past the title screen runs until a valid username/password is
  entered.
- Sidebar now shows "Logged in as `<name>`" with a **Log out** button.
- CSV/Excel/PDF exports now include a "Prepared by: `<student name>`" line.
- Replaced the Gemini-API-key AI feature with the bring-your-own-chat
  workflow described above — no `google-genai` dependency, no API key
  fields anywhere in the app.
- Added a new **🧪 AI Calc Challenge** tab with two rounds, a
  "predict-before-you-look" gate and four-way comparison in Round 2,
  AI-tool + timestamp disclosure on both rounds, per-ratio comparison
  charts, an optional Google Sheets autosave so students can resume across
  logins, and a CSV/PDF export for submitting Round 2 (see above) — purely
  additive, doesn't change any existing calculation, tab, or export.
- Everything else (ratio calculations, yfinance fetch logic, benchmarks,
  existing tabs, downloads) is unchanged from your original file.
