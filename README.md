# VS PHE Swim Rubric

A web app for PE teachers at Chadwick International to assess students against the school's swimming continuum, generate parent-facing comments with AI, and manage multiple spreadsheets of student data across school years — all backed by Google Sheets.

Built as a standalone Google Apps Script web app. One deployment serves every teacher in the domain; no individual logins, no infrastructure to manage.

## Features

- **Grade 1–5 swim rubrics** with Freestyle, Backstroke, Breaststroke, and stamina scoring
- **Dual rubric types** — Standard (level 1–5) and Low-Level (skill checklist) — switchable per student
- **Edit mode** with inline score selection, teacher notes per cell, and auto-save
- **AI-drafted parent comments** via Gemini, calibrated to absolute skill level (not just grade expectation)
- **Multi-spreadsheet support** — one app serves many years of data; teachers switch via a dropdown
- **One-click spreadsheet creation** — clones a template into a designated Drive folder with Grade 1–5 tabs pre-built, blank rows ready for the new class
- **Year-range picker** for the new-spreadsheet flow (displays `2026-2027`, files named `2026-27 VS Swim Rubric`)
- **Print / Save as PDF** via the browser's native print dialog (vector text, A4 landscape)
- **Domain-scoped access** — anyone at Chadwick can use it; no external sign-ups
- **Soft-delete safe** — trashed spreadsheets hide from the dropdown but reappear if restored

## How it works

The app runs on Google Apps Script with `executeAs: USER_DEPLOYING`. That means every action — reading sheets, creating Drive files, calling Gemini — runs as the owner's Google account, so teachers piggyback on one authorization instead of each granting scopes themselves.

```
┌─────────────────┐      ┌──────────────────┐      ┌───────────────────┐
│  Teacher        │──────│  Apps Script     │──────│  Google Sheets    │
│  (browser)      │      │  web app (/exec) │      │  (one per year)   │
└─────────────────┘      │                  │      └───────────────────┘
                         │  doGet / doPost  │      ┌───────────────────┐
                         │  + Gemini API    │──────│  Gemini Flash     │
                         │  + Drive API     │      └───────────────────┘
                         └──────────────────┘      ┌───────────────────┐
                                                   │  Drive folder     │
                                                   │  (new spreadsheet │
                                                   │   copies land)    │
                                                   └───────────────────┘
```

## Repository layout

```
.
├── Code.gs          # All server-side logic: sheet I/O, AI prompt, Drive ops, config
├── Index.html       # Sidebar layout + <?= ?> scriptlets for template data
├── CSS.html         # Flat modern styling (no gradients)
├── JS.html          # Client-side rendering, edit-mode handlers, save queue
├── appsscript.json  # Manifest — scopes, web app config, advanced services
└── README.md        # You are here
```

## Setup

### 1. Prerequisites

- A Google Workspace account with permission to create Apps Script projects
- A Google Sheet with the expected column layout (First Name, Last Name, Class, Teacher, Free, Back, Breast, Stamina, Stamina Kickboard, Additional Comments, Use Low Level Rubric, and the face/bubbles/bob/float/flutter skill columns)
- A [Gemini API key](https://aistudio.google.com/apikey)
- A dedicated Drive folder for new spreadsheet copies

### 2. Create the Apps Script project

1. Go to [script.google.com](https://script.google.com) → **New project**
2. Create the four files (`Code.gs`, `Index.html`, `CSS.html`, `JS.html`) and paste in the contents from this repo
3. Project Settings → **Show "appsscript.json" manifest file in editor**
4. Replace the manifest with the contents of this repo's `appsscript.json`

### 3. Script Properties

Project Settings → Script Properties → add the following five keys:

| Key | Value | Example |
|---|---|---|
| `GEMINI_API_KEY` | Your Gemini API key | `AIzaSy...` |
| `SPREADSHEET_CONFIG` | JSON map of friendly label → spreadsheet ID | `{"2025-2026":"1AbC...xyz"}` |
| `TEMPLATE_SPREADSHEET_ID` | ID of the master template (usually your first/oldest spreadsheet) | `1AbC...xyz` |
| `TARGET_FOLDER_ID` | Drive folder ID where `+ New` copies should land | `1D4e...789` |
| — | — | — |

Spreadsheet and folder IDs come from the URL: `docs.google.com/spreadsheets/d/<THIS_PART>/edit` or `drive.google.com/drive/folders/<THIS_PART>`.

### 4. Deploy

1. **Deploy → New deployment → Web app**
2. Execute as: **User deploying** (this is what lets teachers piggyback on your auth)
3. Who has access: **Anyone within <your domain>**
4. Deploy

You'll be prompted to approve OAuth scopes including full Drive access. Approve them.

### 5. Share the folder with your teachers

Open the target Drive folder → Share → add the teacher group with **Editor** access. Newly created spreadsheets inherit this automatically, so teachers can open them directly when they click **Spreadsheet** in the sidebar.

## Usage

### For teachers

1. Open the `/exec` URL (bookmark it)
2. Pick the spreadsheet for the current school year
3. Filter by class / teacher / student
4. In **View** mode: read rubrics and print/save as PDF
5. In **Edit** mode: click rubric cells to set levels, click the comment icons to add notes, click **AI Comment** to draft a parent-facing comment

### For the owner (creating next year's spreadsheet)

1. Click **+ New** next to the Spreadsheet dropdown
2. Adjust the year range with the ▲/▼ arrows (defaults to the upcoming academic year)
3. Click **Create** — you'll see a confirmation with a link to open the new (empty) spreadsheet
4. The new spreadsheet appears at the top of the dropdown and becomes the default on next load

## Operational notes

- **After any manifest change**, push a new deployment version (Deploy → Manage deployments → pencil → New version). The `/exec` URL doesn't pick up scope changes until you redeploy.
- **Trashing a spreadsheet in Drive** hides it from the dropdown but keeps the entry in `SPREADSHEET_CONFIG`. Restoring from Trash brings it back. Hard-deleted files are hidden permanently; re-create via **+ New**.
- **Duplicate labels** are blocked if the existing file is live, but silently overwritten if the existing file is trashed/deleted (so you can re-create a `2025-2026` after trashing a bad copy).
- **Ownership transfer**: if you ever hand this off, the new owner must re-authorize (trigger any Drive-touching function from the editor once) and push a new deployment version. Teachers do nothing.


## Tech

- **Runtime**: Google Apps Script V8
- **Frontend**: Vanilla JS + HTML + CSS (no build step, no framework)
- **AI**: Gemini Flash Latest via `UrlFetchApp`
- **Storage**: Google Sheets (one per year)
- **Config**: `PropertiesService` script properties


