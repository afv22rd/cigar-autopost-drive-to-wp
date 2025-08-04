# CigarAutopost – Google Drive ➜ WordPress Automation

## Overview

`CigarAutopost` is a **fully-interactive CLI tool** that transforms Google Drive content into fully-formatted WordPress posts. It scans a Google Sheet for rows marked as “Ready To Post”, pulls the associated Google Docs, optionally uploads a featured image from Google Drive, and finally publishes (or drafts) the post on WordPress. Every step is confirmed through a keyboard-driven interface so editors always stay in control.

## Why another README?

The project evolved from a single monolithic script to a modular architecture (`google_integration.py`, `wordpress_integration.py`, etc.). This README was rewritten to describe the current code base, required environment variables and the interactive workflow provided by `main.py`.

## Features

* Detect eligible rows in Google Sheets (✅ Ready To Post, ❌ Online).
* Parse Google Docs for redaction, headline options, cutlines, authors & categories.
* Interactive terminal wizard to pick headline & cutline and decide to:
  * Publish (ENTER)
  * Save as draft (BACKSPACE)
  * Skip (SPACEBAR)
  * Abort (ESC)
* Upload featured images from Google Drive with automatic format validation and retry logic.
* Create missing authors & categories in WordPress on-the-fly.
* Update the Google Sheet once the post goes online.

## Folder & Module layout

| File | Responsibility |
|------|----------------|
| `constants.py` | Loads environment variables, initialises Google API clients, defines ANSI colours. |
| `google_integration.py` | Reads & updates Google Sheets, parses Google Docs (redaction, headlines, cutlines). |
| `image_processing.py` | Downloads images from Drive and uploads them to the WordPress media library with retry logic. |
| `wordpress_integration.py` | Creates authors, categories and posts through the WordPress REST API. |
| `user_interface.py` | Handles all terminal I/O – single-key reading and the rich post review screen. |
| `main.py` | Orchestrates the complete flow; entry point executed by the user. |

## Requirements

```bash
python >= 3.9
pip install -r requirements.txt
```

Key packages pinned in `requirements.txt` include (but are not limited to):

```
google-api-python-client
google-auth
python-dotenv
requests
```

## Configuration (.env)

Create a `.env` file in the project root:

```
GOOGLE_CREDENTIALS_FILE=/absolute/path/to/service_account.json
WP_URL=https://your-site.com/wp-json
WP_USER=your_wp_username
WP_PASSWORD=application_password_generated_in_wp
```

The credentials file must belong to a **service account** that has access to the target Drive, Docs and Sheets.

## Google Cloud set-up (one-time)

1. Create a Google Cloud project & service account.
2. Enable the following APIs:
   * Google Sheets API
   * Google Drive API
   * Google Docs API
3. Share the target Drive folder and the Sheet with the service-account e-mail (viewer/editor access).

## WordPress set-up (one-time)

1. Log in to WordPress as an administrator.
2. Go to **Users ▸ Profile ▸ Application Passwords** and generate a new token.
3. Copy the value into `WP_PASSWORD` in your `.env`.

## Spreadsheet layout expected by the script

| Column | Purpose |
|--------|---------|
| **B** | Ready To Post (checkbox) |
| **D** | Online (checkbox) – will be set automatically |
| **E** | Story: link to Google Doc containing the redaction |
| **H** | Author(s) (comma-separated) |
| **N** | Featured image: Google Drive link |
| **O** | Categories (comma-separated) |
| **P** | Headlines document (optional) |
| **Q** | Cutlines document (optional) |

Section headers are read from column **A**. When no categories are supplied, the current section name is used as a fallback.

## Google Doc structure

```
Headline:
Featured image:
Cutlines:
Redaction:
Author(s):
Categories:
```

Only *Headline* and *Redaction* are strictly required; any missing fields will be prompted for interactively.

## Running

```bash
$ python main.py
Enter Google Sheets URL: https://docs.google.com/spreadsheets/d/•••
```

After the Sheet ID is validated the interactive wizard starts.

### Keyboard shortcuts inside the wizard

| Key | Action |
|-----|--------|
| **ENTER** | Publish the post immediately |
| **BACKSPACE** | Save the post as a draft |
| **SPACEBAR** | Skip this row |
| **ESC** | Terminate the program |

## Logs

All operations are printed to stdout with colour coding (green = success, yellow = warning, red = error). At the end of a run a per-section and overall summary is displayed.

## Limitations

* Only one featured image per post is supported.
* Advanced WordPress blocks/layouts are not generated – the output is simple HTML paragraphs.
* The tool requires a TTY for interactive controls; it will not run inside non-interactive environments like cron without modification.

## License

MIT
