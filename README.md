# MiiR Routing Assistant

Self-serve decorator routing tool for MiiR reps. Pick a silhouette (and size/variant), and it shows a capability matrix across every decoration type and decorator, lead times with business-day ship dates, and eligible decorators — sourced from the MiiR Printer Availability Template.

## Quick start

### 1. Install dependencies

```bash
pip3 install -r requirements.txt
```

### 2. Verify your environment

```bash
python3 verify_template.py
```

This confirms pandas/openpyxl/streamlit/pgeocode are installed and the template is in the right place.

### 3. Run the app

```bash
streamlit run app.py
```

Opens at [http://localhost:8501](http://localhost:8501).

## Deployed version

The live version of this tool runs on [Streamlit Community Cloud](https://streamlit.io/cloud). Use that link for day-to-day use — only run it locally for development or template testing.

## Updating the template

1. Edit `MiiR_Printer_Availability_Template_v6.xlsx` in place.
2. The app reloads automatically — no restart needed. Data is cached for 60 seconds; wait one minute and refresh to see changes.

The footer shows the file's last-modified timestamp so you always know which version is loaded.

## Template sheets used

| Sheet | Purpose |
|---|---|
| HOLIDAYS | MiiR-observed holidays excluded (along with weekends) from ship-date math |
| PRODUCTS | Silhouette (Family) and Size/Variant dropdown options |
| DECORATORS | Decorator profiles (region, cost tier, zip code for proximity sort) |
| DECORATION_TYPES | Lead times per decoration type |
| CAPABILITY | Product × Decoration Type × Decorator capability matrix |
| PREFERRED_VENDOR | Present in template for Phase 2 — not wired into the app |
| CUSTOMER_PREFERENCES | Present in template for Phase 2 — not wired into the app |

## Capability values

| Value | Meaning | Matrix color |
|---|---|---|
| YES | Fully capable | Green |
| YES* | Capable with conditions | Yellow |
| LIMITED | Partial capability — check caveats | Orange |
| UNTESTED | Not yet evaluated | Grey |
| NO | Cannot decorate this combination | Red |

MPIX is a backup-only decorator kept in the DECORATORS sheet for record-keeping, but it never appears in the app's matrix or eligible-decorator lists.

## Zip-code proximity

Enter the customer's shipping zip to sort eligible decorators by driving distance (great-circle, via [pgeocode](https://pypi.org/project/pgeocode/)) instead of capability tier. If a decorator's zip can't be geocoded, it's shown as "distance unknown" at the bottom of the list rather than breaking the sort.

## What's deferred (Phase 2)

- Sales order CSV parsing
- JOB_RULES validation
- Inventory pre-check
- Acumatica / Google Drive integration
- PREFERRED_VENDOR / CUSTOMER_PREFERENCES routing overrides
