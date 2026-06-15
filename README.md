# Concept Engineers — Fee Proposal Generator

A web-based tool that lets staff fill in a tick-box form and download a finished, branded fee proposal as a PDF or Word document in under two minutes.

---

## What it does

Staff open a local web page, tick the relevant scope items, fill in client details and fees, and click **Generate Proposal**. The app renders the master Word template with those inputs and returns a ready-to-send PDF (or DOCX) for download.

No manual formatting. No copy-pasting between documents. No blank pages from unused sections.

---

## What's been built so far

| File | Purpose |
|---|---|
| `master_proposal_template.docx` | The branded Word template the engine fills. Contains all conditional logic via docxtpl/Jinja2 tags. |
| `build_template.py` | Regenerates `master_proposal_template.docx` from scratch. Edit this file to change layout, sections, or branding. |
| `render_example.py` | Standalone script — renders three example proposals (Scenarios A, B, C) directly to DOCX. Useful for testing the template without the web UI. |
| `app.py` | Flask web server. Serves the form and handles the generate endpoint. |
| `templates/form.html` | The single-page web form with all tick boxes, fields, and dynamic rows. |
| `TEMPLATE_VARIABLES.md` | Data contract — every variable the template expects, with types and examples. Read this before touching the template. |
| `sample_A_da_dd.docx` | Example output: DA + DD scope, residential capability, with hourly rates. |
| `sample_B_da_only.docx` | Example output: DA only, commercial capability. |
| `sample_C_uu_only.docx` | Example output: UU works only, stormwater capability. |

### What the form covers

- **Document meta** — proposal ref, date, revision
- **Client & project** — contact, company, site address, email, project understanding paragraph
- **Scope tick boxes** — DA, DD, Urban Utilities (UU), Certifier/UW (with Minor/Major selector), UU Works Only
- **Capability profile** — four built-in tiles (Residential, Commercial, Infrastructure, Stormwater) plus a Custom option
- **Fee schedule** — dynamic rows (add/remove); optional hourly rates table
- **Assumptions & exclusions** — dynamic bullet lists
- **Sender details** — pre-filled with Gina Harper's defaults; collapsible section

### What the template covers

- Firm letterhead header and page-number footer on every page
- Conditional scope sections — only the ticked items appear; no blank paragraphs left behind
- Fee table with looped rows and a bold total row
- Optional hourly-rates table (hidden entirely when empty)
- Assumptions and exclusions as bullet lists (headings hidden when lists are empty)
- Standard Terms & Conditions hyperlink (baked into the template, not a variable)
- Acceptance / signature block

---

## Requirements

- **Python 3.10+**
- **Microsoft Word** (required for PDF export — `docx2pdf` drives Word via COM on Windows)

---

## Setup

```bash
# 1. Clone the repo
git clone https://github.com/devnomanaslam/concept-engineers-proposal-gen.git
cd concept-engineers-proposal-gen

# 2. Install dependencies
pip install flask docxtpl docx2pdf

# 3. Start the server
py app.py
```

Then open **http://localhost:5000** in your browser.

> **PDF fallback:** If Microsoft Word is not installed, clicking Generate will return a DOCX instead of a PDF. No error is shown to the user — the file just downloads as `.docx`.

---

## Running without the web UI (template testing only)

```bash
py render_example.py
```

This renders three scenarios directly to DOCX files in the project folder. Useful for checking template changes without starting the server.

To regenerate the master template after editing `build_template.py`:

```bash
py build_template.py
```

---

## How the template engine works

The template uses [docxtpl](https://docxtpl.readthedocs.io/) (a Jinja2 wrapper for python-docx). Three tag types do all the conditional logic:

| Tag type | Effect |
|---|---|
| `{{ variable }}` | Inline value substitution |
| `{%p if flag %} … {%p endif %}` | The `p` prefix removes the **whole paragraph**, so hidden sections leave no blank lines |
| `{%tr for x in list %} … {%tr endfor %}` | Placed on their own rows, removes/repeats the entire table row |

**Important:** Never put `{%tr for %}` and `{%tr endfor %}` in the same table row — docxtpl will drop the opening tag.

Full variable reference: see [`TEMPLATE_VARIABLES.md`](TEMPLATE_VARIABLES.md).

---

## What still needs to be done

These are the next logical steps — the core "under 2 minutes" proof of concept is complete.

### High priority

- [ ] **Remaining capability profiles** — 5 of 9 are stubbed out. Add: `townhouse`, `mixeduse`, `civil_works`, `land_development`, `civil_certification`. Copy any existing `{%p if capability_profile == 'key' %}` block in `build_template.py`, add the key to `KNOWN_PROFILES` in both `app.py` and `render_example.py`, then run `py build_template.py` to regenerate the template.
- [ ] **Firm ABN and contact details** — currently placeholder values (`00 000 000 000`, `+61 7 0000 0000`). Update the defaults in `app.py` → `build_context()` and in `build_template.py` header.
- [ ] **Real terms URL** — the hyperlink in `build_template.py` → `add_hyperlink(...)` points to `conceptengineers.com.au/terms`. Confirm this page exists and is correct.

### Medium priority

- [ ] **Email delivery** — instead of (or in addition to) downloading the PDF, send it directly to the client email address captured in the form. Use `smtplib` or an API like SendGrid/Mailgun.
- [ ] **Proposal numbering** — auto-increment the proposal ref (e.g. read the last ref from a JSON file and increment). Currently staff type it manually.
- [ ] **Logo in letterhead** — `build_template.py` writes a text-only header. Replace with an image using `doc.add_picture()` or insert the logo into the header paragraph.
- [ ] **Hosted deployment** — move off `localhost`. Options: a small VPS running the Flask app behind nginx, or a PaaS like Railway/Render. Note: PDF export requires a Windows host with Word, or swap `docx2pdf` for LibreOffice headless on Linux.

### Lower priority / future

- [ ] **e-Signing** — integrate DocuSign or Adobe Sign API. Send the PDF for signing directly from the generate endpoint.
- [ ] **HubSpot sync** — POST proposal metadata (ref, client, fees, date) to HubSpot via their CRM API after generation.
- [ ] **Reminder emails** — if no signed copy received in N days, send a follow-up. Needs a job scheduler (APScheduler or a cron task).
- [ ] **Proposal archive** — save each generated proposal (PDF + context JSON) to a folder or database for records. Currently nothing is persisted after download.
- [ ] **Auth** — the app is currently open with no login. Add basic HTTP auth or a simple password gate before putting it on a shared network.

---

## Project contacts

| Role | Name |
|---|---|
| Client / owner | Concept Engineers |
| Original build | Claude Code (AI) + Noman Aslam |
