# Screenshot to Contacts

Extract HR email addresses, phone numbers, company names, and job profiles from screenshot images using a vision-capable AI model. Results are grouped by company and written to CSV and Excel files.

## Setup

Requires Python 3.10 or newer.

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
Copy-Item .env.example .env
```

Add an API key to `.env`. Use either Anthropic directly or OpenRouter through the OpenAI-compatible API.

## Usage

Put screenshots in a folder such as `hr_shots` and run:

```powershell
python .\screenshot-to-contacts.py --provider openai --model anthropic/claude-haiku-4.5 --input .\hr_shots
```

Useful options:

- `--test` checks the provider connection without processing images.
- `--limit N` processes only the first `N` new screenshots.
- `--output NAME` changes the output prefix; the default is `hr_contacts`.
- `--provider anthropic` uses Anthropic directly and the provider's default model.

The script creates `NAME.csv`, `NAME.xlsx`, and `NAME_processed.txt`. These generated files are ignored by Git, and processed screenshots are remembered so interrupted runs can be resumed.

## Notes

- Never commit `.env` or API keys.
- Review extracted contacts before using them; OCR and model output can be imperfect.
