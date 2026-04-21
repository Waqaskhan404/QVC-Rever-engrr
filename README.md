# QVC Appointment Bot

Automated Qatar Visa Center appointment booking and rescheduling bot.

---

## Setup After Cloning

### 1. Create Virtual Environment
```bash
python -m venv venv312
venv312/Scripts/pip install -r requirements.txt
```

### 2. Add Missing Files Manually

These files are **not on GitHub** (they contain sensitive credentials). Copy them into the project folder manually:

| File | How to get it |
|---|---|
| `bot-key.pem` | Copy from original machine (USB / Google Drive / OneDrive) |
| `Webshare residential proxies.txt` | Copy from original machine (USB / Google Drive / OneDrive) |

> Never upload these files to GitHub.

---

## Scripts

| Script | Purpose |
|---|---|
| `qvc_book_api.py` | Book a new appointment |
| `qvc_direct_api.py` | Reschedule an existing appointment |

### Run Booking Bot
```bash
venv312/Scripts/python qvc_book_api.py
```

### Run Reschedule Bot
```bash
venv312/Scripts/python qvc_direct_api.py
```

---

## Configuration

Edit the `CONFIG` section at the top of each script before running:

```python
PASSPORT_NUMBER     = "your passport number"
VISA_NUMBER         = "your visa number"
MOBILE_NUMBER       = "your mobile number"
EMAIL               = "your email"
QVC_CENTER          = "Islamabad"   # or "Karachi"
MONTHS_TO_CHECK     = ["April", "May"]
URGENT_MEDICAL_DATE = "2026-05-05"  # only book before this date, set "" for any
POLL_INTERVAL       = 5             # seconds between scans
DRY_RUN             = False         # True = test mode, no actual booking
```

---

## Alerts

Discord alerts fire automatically when slots are found:

- **Urgent slots** (before `URGENT_MEDICAL_DATE`) → alert on **Urgent webhook**
- **Normal slots** → alert on **Normal webhook**
- If running in urgent mode and normal slots found → alert on normal webhook, **not booked**







(Set-ExecutionPolicy -Scope Process -ExecutionPolicy RemoteSigned) ; (& c:\Users\waqas\Desktop\QVC_Production\venv312\Scripts\Activate.ps1)  