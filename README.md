# HPI SmartWe Calendar Export

Export your calendar events from HPI SmartWe portal to iCal format.

## Features

- Automated login through Microsoft SAML, HPI ADFS, and SmartWe portal
- Extracts events from "Meine Termine" (My Appointments)
- Exports to standard `.ics` format for import into any calendar app

## Setup

### 1. Install dependencies

```bash
pip install -r requirements.txt
playwright install chromium
```

### 2. Create credentials file

Create a `.credentials` file in the project directory with your login details:

```
page: login.microsoftonline.com
user: your.email@student.hpi.uni-potsdam.de
pw: your_microsoft_password

page: adfs.hpi.uni-potsdam.de
user: your.username
pw: your_hpi_password

page: sv-portal.hpi.de
user: your.username
pw: your_smartwe_password
```

**Note:** The `.credentials` file is in `.gitignore` and will not be committed.

## Usage

```bash
python extract_calendar.py
```

The script will:
1. Open a browser window
2. Automatically log in through all three authentication steps
3. Navigate to "Termine" → "Meine Termine"
4. Extract all your calendar events
5. Save them to `calendar_export.ics`

## Output

The exported `calendar_export.ics` file can be imported into:
- Google Calendar
- Apple Calendar
- Outlook
- Any other calendar app supporting iCal format

## Troubleshooting

### Login fails
- Check your credentials in `.credentials`
- Make sure you have the correct passwords for all three login steps
- The SmartWe portal may have different credentials than ADFS

### Not all events exported
- The SmartWe interface uses a virtualized table
- Currently exports visible events; scroll handling is WIP

## Files

- `extract_calendar.py` - Main script
- `requirements.txt` - Python dependencies
- `.credentials` - Your login credentials (create this yourself)
- `calendar_export.ics` - Output file (generated)

## License

MIT
