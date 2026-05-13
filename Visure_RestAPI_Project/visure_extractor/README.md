# Visure API → Excel Extractor

A small Python tool that pulls requirements data from Visure Authoring via its
REST API and writes a structured Excel workbook. The Excel file is the data
source for a Power BI dashboard.

**Author:** Prince Punshi (Senior Rail Engineer, Mott MacDonald)
**Status:** v0.1 — single-project interactive mode.

---

## What it does

1. Logs into your Visure instance using credentials from a `.env` file.
2. Shows you the projects your account can see, lets you pick one.
3. Lists the specifications in that project, lets you pick one or all.
4. For each chosen spec, pulls every element with all user attributes
   in a single API call.
5. Writes a three-sheet `.xlsx`:
   - **Requirements** — one row per element, with standard columns plus
     one column per discovered attribute.
   - **Summary** — one row per spec showing element counts and status.
   - **Run Log** — extraction timestamps, duration, any errors.

The Excel goes into `./output/` with a timestamped filename so successive
runs don't overwrite each other.

---

## Setup (one time, ~5 minutes)

### 1. Install Python 3.10+

```bash
python --version
```

If that prints `Python 3.10.x` or higher, you're good. Otherwise download
from [python.org](https://www.python.org/downloads/) and tick "Add Python to PATH"
during install.

### 2. Install the dependencies

From inside the project folder (the one containing `run.py`):

```bash
pip install -r requirements.txt
```

If you prefer a virtual environment (recommended — keeps these packages
separate from anything else on your machine):

```bash
python -m venv .venv
.venv\Scripts\activate          # Windows
# OR
source .venv/bin/activate        # macOS / Linux
pip install -r requirements.txt
```

### 3. Create your `.env` file

```bash
copy .env.example .env           # Windows
# OR
cp .env.example .env             # macOS / Linux
```

Open `.env` in any text editor and fill in:

```
VISURE_USERNAME=your.username
VISURE_PASSWORD=your_password_here
VISURE_BASE_URL=https://mottmac.visurecloud.com/VisureAuthoring8
```

The `.env` file is already in `.gitignore` so it will never be committed
to version control. **Do not paste it into emails or chat.**

---

## Running it

```bash
python run.py
```

You'll see something like:

```
============================================================
Visure API → Excel extractor
Base URL: https://mottmac.visurecloud.com/VisureAuthoring8
User:     prince.punshi
============================================================

[1/4] Authenticating...
      OK

[2/4] Available projects:
---------------------------
  [ 1] T3PP Brisbane Airport  (id=42)
  [ 2] Some Other Project     (id=99)
Pick a project (1-2): 1
      Selected: T3PP Brisbane Airport

[3/4] Loading specifications for 'T3PP Brisbane Airport'...
      Found 12 specification(s).

Do you want to extract:
  [1] All specifications in this project
  [2] A single specification
Choice (1-2): 1
      Will extract 12 specification(s).

[4/4] Extracting elements + attributes...
      (1/12) System Requirements ... 247 elements
      (2/12) Interface Requirements ... 118 elements
      ...

Writing Excel workbook...
      Wrote: C:\...\output\visure_extract_20260514_103247.xlsx

Done. 12/12 specs OK, 1,842 elements total.
```

---

## How it's structured

```
visure_extractor/
├── .env.example       ← template (safe to commit)
├── .env               ← your real credentials (gitignored)
├── .gitignore
├── requirements.txt   ← Python dependencies
├── run.py             ← entry point
├── README.md          ← this file
└── visure/
    ├── __init__.py
    ├── config.py      ← loads .env into a Settings object
    ├── client.py      ← all Visure API calls live here
    ├── extractor.py   ← orchestration + user prompts
    └── excel_writer.py← flattens elements → DataFrame → Excel
```

**Why split into files?** So you can change one piece without breaking the
others. For example, swapping from Excel to direct-to-Parquet output is a
change to `excel_writer.py` only. Adding a new API endpoint is a change
to `client.py` only.

---

## Key design decisions (and why)

### Read-only by construction
The client only ever calls `GET` endpoints, with these exceptions:
- `POST /api/v1/authenticate` — required to get a token.
- `POST /api/v1/project/current` — required by Visure to pin the session
  to a project before listing its specs.
- `POST /api/v1/logout` — best-practice cleanup at the end.

There are **no** POST/PUT/DELETE calls against requirements, attributes,
or any data. The script is physically incapable of modifying Visure data.
For an extra layer of safety, ask your Visure admin to create a dedicated
read-only service account and put its credentials in `.env`.

### One API call per spec for everything
The endpoint `GET /api/v1/specification/{id}/items?includeAllAttributes=true`
returns the entire element hierarchy with every attribute baked in. We don't
loop per-element to fetch attributes — that would be hundreds of calls.
One project = roughly N+3 API calls where N is the spec count.

### Re-authenticate fresh on every run
No persistent token cache, no heartbeat. Each run starts clean. Power BI
schedules the run; the run is short (seconds for one project); token
expiry is irrelevant.

### Session quirk
Visure's REST API has a stateful concept of "current project". Calling
`GET /api/v1/specifications` returns specs **for whichever project was last
selected with `POST /project/current`**. The `client.set_current_project()`
method hides this — but it's worth knowing about if you extend the script.

---

## Configuring Power BI

1. In Power BI Desktop: **Get Data → Excel workbook**.
2. Point it at the most recent file in `output/`.
3. Select the **Requirements** sheet.
4. Build whatever visuals you want.

To make it pick up new extractions automatically, point Power BI at a fixed
filename (e.g. `output/visure_latest.xlsx`) and modify `excel_writer.py` to
write to that name as well as the timestamped one. We can do that in v0.2.

---

## Troubleshooting

**`Authenticated, but no token in response`**
The auth call succeeded but the token wasn't where we expected. Check that
your account has `AUTHORING` license type, or change `LICENSE_TYPE` in `.env`.

**`401 Unauthorized` mid-run**
Token expired. Unlikely for a single-project run that finishes in seconds.
If you hit this with the "all projects" mode in a future version, we add
re-authentication between projects (already designed for, not yet wired up).

**`This project has no specifications`**
Either the project really has none, or your account doesn't have access.
Try it in the Visure web UI to confirm.

**Power BI says "the file is open in another program"**
Close the Excel file before re-running the extractor.

---

## Roadmap

- v0.2 — "all projects" mode with re-auth between projects.
- v0.3 — Optional traceability links sheet (`/element/{id}/relationships`).
- v0.4 — Direct Power BI Python connector (skip the Excel intermediate).
- v0.5 — Service-account / Windows Credential Manager support.
