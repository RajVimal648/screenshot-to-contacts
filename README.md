# Screenshot to Contacts

Extract HR and recruiter contact details from screenshots. The tool reads job posts, chat messages, and listings from images using a vision-capable AI model, then saves the email addresses, phone numbers, company names, and job profiles to CSV and Excel, grouped by company.

## Contents

- [Features](#features)
- [Requirements](#requirements)
- [Installation](#installation)
- [Configuration](#configuration)
- [Usage](#usage)
- [Output](#output)
- [Resuming and re-running](#resuming-and-re-running)
- [Troubleshooting](#troubleshooting)
- [Privacy and responsible use](#privacy-and-responsible-use)
- [Repository layout](#repository-layout)

## Features

- Reads `.png`, `.jpg`, `.jpeg`, and `.webp` screenshots, including LinkedIn, Naukri, and WhatsApp images.
- Extracts email addresses, phone numbers, company name, job profile, and recruiter name.
- Writes **one row per company**. Multiple emails or phone numbers for the same company are joined with commas.
- Normalises and de-duplicates results across all screenshots.
- Supports Anthropic directly, or any OpenAI-compatible API such as OpenRouter.
- Resumes safely: processed screenshots are remembered, and failed ones are retried on the next run.
- Produces both a `.csv` and a formatted `.xlsx` file.

## Requirements

- Python 3.10 or newer
- An API key for one of the supported providers (see [Configuration](#configuration))
- A vision-capable model. The examples use Claude Haiku 4.5, which is inexpensive and reads screenshot text well.

## Installation

**Windows (PowerShell)**

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
Copy-Item .env.example .env
```

**macOS / Linux**

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
cp .env.example .env
```

`requirements.txt` should contain:

```text
anthropic
openai
pillow
openpyxl
python-dotenv
```

## Configuration

Edit `.env` and set the variables for the provider you use. Values in `.env` take precedence over variables already set in your terminal session. Do not wrap values in quotes.

### OpenRouter or another OpenAI-compatible API

```dotenv
OPENAI_BASE_URL=https://openrouter.ai/api/v1
OPENAI_API_KEY=your-api-key
```

The base URL must end in `/v1` (or the path your provider documents for its OpenAI-compatible endpoint).

### Anthropic

```dotenv
ANTHROPIC_API_KEY=your-api-key
```

`ANTHROPIC_AUTH_TOKEN` is also accepted. Set `ANTHROPIC_BASE_URL` only if you route requests through a gateway.

> **Never commit `.env` or API keys.** Add `.env` to `.gitignore`, and rotate any key that has been shared or displayed in a screenshot.

## Usage

Place your screenshots in a folder, for example `hr_shots`, then work through these three steps.

**1. Check the connection** (sends a short text message; no images are processed):

```powershell
python .\screenshot-to-contacts.py --provider openai --model anthropic/claude-haiku-4.5 --test
```

**2. Try a small batch** and review the result:

```powershell
python .\screenshot-to-contacts.py --provider openai --model anthropic/claude-haiku-4.5 --input .\hr_shots --limit 3
```

**3. Process the whole folder:**

```powershell
python .\screenshot-to-contacts.py --provider openai --model anthropic/claude-haiku-4.5 --input .\hr_shots
```

### Options

| Option | Description | Default |
| --- | --- | --- |
| `--provider {anthropic,openai}` | API format to use. `openai` covers OpenRouter and other OpenAI-compatible services. | `anthropic` |
| `--model ID` | Model ID as named by your provider. | `claude-haiku-4-5-20251001` (anthropic), `anthropic/claude-haiku-4.5` (openai) |
| `--input DIR` | Folder containing screenshots. Required unless `--test` is used. | none |
| `--output NAME` | Output file prefix. | `hr_contacts` |
| `--limit N` | Process only the first `N` new screenshots. | all |
| `--test` | Verify the provider connection and exit. | off |

## Output

Each run creates or updates three files in the working directory:

| File | Purpose |
| --- | --- |
| `NAME.csv` | Extracted contacts, one row per company |
| `NAME.xlsx` | The same data, formatted for Excel |
| `NAME_processed.txt` | List of screenshots already processed |

### Columns

| Column | Content |
| --- | --- |
| `company` | Hiring company name |
| `profile` | Job title or role. Multiple values are separated by ` \| `. |
| `emails` | Email addresses, separated by commas |
| `phones` | Phone numbers, separated by commas |
| `hr_name` | Recruiter or poster name. Multiple values are separated by ` \| `. |
| `source_file` | Screenshot(s) the row came from |

### Example

| company | profile | emails | phones | hr_name | source_file |
| --- | --- | --- | --- | --- | --- |
| Acme Pvt Ltd | .NET Developer \| QA Engineer | hr@acme.com, jobs@acme.com | +919876543210, 9123456789 | Priya | shot1.jpg \| shot4.jpg |
| Zeta | Backend Developer | jobs@zeta.in | | Ravi | shot2.jpg |

### How records are cleaned and merged

- **Company matching** ignores case, punctuation, and common legal suffixes (`Pvt`, `Ltd`, `Private`, `Limited`, `LLP`, `Inc`, `Corp`, `Corporation`). "Acme Pvt. Ltd." and "ACME" are treated as the same company.
- **Entries without a company name** are kept as separate rows, since there is nothing to group them by.
- **Emails** are lower-cased and validated. Obfuscated forms such as `hr [at] abc [dot] com` are converted to normal addresses.
- **Phone numbers** are reduced to digits, keeping a leading `+`, and must contain 8 to 15 digits. A number with and without its country code is treated as the same number.
- **Duplicates** are removed across all screenshots.

> Open the `.xlsx` file rather than the `.csv` in Excel. Excel can convert a phone number in a CSV to scientific notation and drop the leading `+`. The `.xlsx` stores phone numbers as text.

## Resuming and re-running

- Screenshots that finish successfully are recorded in `NAME_processed.txt` and skipped on later runs, so you can add new screenshots to the folder and run the same command again.
- Screenshots that fail are not recorded and are retried on the next run.
- The run stops after three consecutive failures, so a configuration problem does not consume credits on every image.
- Existing output is loaded and merged on each run. A CSV in the older one-row-per-email format is merged into the current layout automatically.
- To reprocess screenshots, delete `NAME_processed.txt`. Emails and phone numbers already saved are not duplicated.
- Close the `.csv` and `.xlsx` files in Excel before running. Windows locks open files and the script cannot save over them.

## Troubleshooting

| Message | Likely cause and fix |
| --- | --- |
| `API_KEY not found` | The key is missing from `.env`, or `.env` is not in the project folder. |
| `401` invalid or missing key | The key is wrong, incomplete, or belongs to a different provider or gateway than the base URL. |
| `402` requires more credits | The account balance is too low. Add credits with your provider. |
| `403` no access to model | Your key is not permitted to use that model. Check the model ID and your key settings. |
| `unauthorized client detected` | The gateway only accepts requests from approved applications and rejects custom scripts. Use a provider that allows direct API access. |
| `404 Not Found` | The base URL or model ID is wrong. For OpenAI-compatible providers the base URL usually ends in `/v1`. |
| `'str' object has no attribute 'choices'` | The base URL points at a website rather than an API endpoint. Check the `/v1` suffix. |
| `Cannot write ... Close them in Excel` | The output files are open in Excel. Close them and re-run. |
| `model did not return JSON` | The model replied in an unexpected format. Retry, or try a stronger model for that screenshot. |

## Privacy and responsible use

- Screenshots are sent to the AI provider you configure, and to any gateway in between. They contain personal data such as names, emails, and phone numbers. Review your provider's data policy before processing sensitive images.
- Extraction is imperfect. Blurry or cropped text can be misread, and a single wrong character makes an email address invalid. Review the output before relying on it.
- Use collected contact details only for legitimate purposes, such as applying for jobs, and comply with the data-protection and anti-spam laws that apply to you.

## Repository layout

```text
.
├── screenshot-to-contacts.py   # main script
├── requirements.txt            # Python dependencies
├── .env.example                # template for API configuration
├── .gitignore
└── README.md
```

Suggested `.gitignore` entries:

```gitignore
.env
.venv/
hr_shots/
hr_contacts.csv
hr_contacts.xlsx
hr_contacts_processed.txt
```