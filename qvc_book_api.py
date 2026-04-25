"""
QVC Direct API — Book New Appointment
======================================
Calls Qatar Visa Center APIs directly (no browser) to book a new appointment.
Uses the real book-appointment flow captured from Chrome:
  populateCaptcha (empty body) → validatevisaandpass → tokenValidation → scan dates → book

Usage:
    venv312/Scripts/python qvc_book_api.py

Adjust CONFIG section below before running.
"""

# ── stdlib ────────────────────────────────────────────────────────────────────
import base64
import calendar
import io
import json
import os
import random
import sys
import time
import uuid
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timedelta

# ── third-party ───────────────────────────────────────────────────────────────
from curl_cffi import requests
from Crypto.Cipher import AES
from Crypto.Util.Padding import pad, unpad
from PIL import Image, ImageFilter, ImageEnhance
import numpy as np
import easyocr

# ── ddddocr ───────────────────────────────────────────────────────────────────
try:
    import ddddocr as _ddddocr
    _DDDD = _ddddocr.DdddOcr(show_ad=False)
    _DDDD_AVAILABLE = True
except Exception:
    _DDDD_AVAILABLE = False

# ── CNN captcha model ─────────────────────────────────────────────────────────
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "captcha_solver"))
try:
    import torch as _torch
    import torchvision.transforms as _T
    from model import CaptchaCNN as _CaptchaCNN, decode_ctc as _decode_ctc
    _CNN_AVAILABLE = True
except Exception:
    _CNN_AVAILABLE = False

# ── CONFIG — edit these before running ───────────────────────────────────────

PASSPORT_NUMBER   = "CF6797681"
VISA_NUMBER       = "382026063699"
MOBILE_NUMBER     = "00923045454166"
EMAIL             = "waqas.khan.40004@gmail.com"

# Which VSC to book appointment AT
QVC_CENTER        = "Islamabad"        # "Islamabad" or "Karachi"

# Months to scan for available slots (order matters)
MONTHS_TO_CHECK   = ["April", "May"]

# Only book a date STRICTLY BEFORE this YYYY-MM-DD date. Set to "" or None to book any date.
URGENT_MEDICAL_DATE = "2026-05-06"

# How many seconds to wait between polling cycles (when no slots found)
POLL_INTERVAL     = 3

# DRY_RUN = True → go through the full flow but skip the final save (for testing)
DRY_RUN = False

# Proxy file — one proxy per line: host:port:user:pass
PROXY_FILE = os.path.join(os.path.dirname(__file__), "Webshare residential proxies.txt")

# ── CONSTANTS ─────────────────────────────────────────────────────────────────

BASE_URL  = "https://agent.qatarvisacenter.com"
ORIGIN    = "https://www.qatarvisacenter.com"

DISCORD_WEBHOOK        = "https://discordapp.com/api/webhooks/1495888024561516777/Hdcv7CY-fE8zjtxo3eYupCVLYzapg_cIlY3lbF0YSLWWyor1TIq7hBWYMCn4RsF2TOGO"
DISCORD_URGENT_WEBHOOK = "https://discord.com/api/webhooks/1496147361569833140/-LydSfbfDWM0KWlEjhePMABhEMzgTbp8i0wm_aiHPJ7565KdIMQMFKMZhgQMd7zZnK0Q"
DISCORD_SERVER_WEBHOOK = "https://discord.com/api/webhooks/1496431421349302454/hGxi17liWF64k-Amp7QzxF9tJm1cWtBwoeV2S0pe49haa5sdf8wurqlYKrvEZ9r6BJA6"

VSC_MAP = {
    "Islamabad": {"vscId": 4050, "vscName": "Islamabad", "vscCode": "IS"},
    "Karachi":   {"vscId": 4051, "vscName": "Karachi",   "vscCode": "KC"},
}

COUNTRY_TO = {
    "countryCode": "PK", "countryDisplayName": "Pakistan",
    "countryId": 3050, "countryName": "Pakistan",
    "languageDisplayName": "English", "languageId": 1, "languageName": "English",
    "currencyCode": "PKR", "isGroupEnabled": "N", "isLoungeEnabled": "N",
    "isQicEnabled": "N", "showLandingPopup": "Y", "amenitiesLoungePopup": "Y",
    "isLoungePopUpEnabled": "Y", "isAppointmentBookingEnabled": "Y",
    "isAldarEnabled": "Y", "alertMessage": {"message": "", "display": False},
}

MONTH_NAMES = {
    "January": 1,  "February": 2,  "March": 3,    "April": 4,
    "May": 5,      "June": 6,      "July": 7,      "August": 8,
    "September": 9,"October": 10,  "November": 11, "December": 12,
}

AES_KEY    = b"cvq@4202temoib!&"
PKT_OFFSET = timedelta(hours=5)

# Placeholder values used by browser for visaNumber/passportNumber in booking calls
_DUMMY_VISA = "123456"
_DUMMY_PASS = "123456"

# ── AES helpers ───────────────────────────────────────────────────────────────

def _aes_encrypt(plain: str) -> str:
    cipher = AES.new(AES_KEY, AES.MODE_CBC, iv=AES_KEY)
    ct = cipher.encrypt(pad(plain.encode("utf-8"), 16))
    return base64.b64encode(ct).decode("ascii")


def _aes_decrypt(b64: str) -> str:
    ct = base64.b64decode(b64)
    cipher = AES.new(AES_KEY, AES.MODE_CBC, iv=AES_KEY)
    return unpad(cipher.decrypt(ct), 16).decode("utf-8")


def _enc(payload: dict) -> dict:
    return {"encryptedData": _aes_encrypt(json.dumps(payload, separators=(",", ":")))}


def _dec(resp: dict) -> dict:
    return json.loads(_aes_decrypt(resp["encryptedData"]))


def _encode_url_safe(value: str) -> str:
    b64 = _aes_encrypt(value)
    return b64.replace("/", "_").replace("+", "-")


# ── Proxy management ──────────────────────────────────────────────────────────

_proxies: list[str] = []
_proxy_idx: int = 0


def _load_proxies():
    global _proxies
    if not PROXY_FILE:
        return
    try:
        with open(PROXY_FILE, "r") as f:
            _proxies = [l.strip() for l in f if l.strip()]
        print(f"[PROXY] Loaded {len(_proxies)} proxies from {PROXY_FILE}")
    except Exception as e:
        print(f"[PROXY] Could not load proxy file: {e}")


def _pick_proxy() -> str | None:
    global _proxy_idx
    if not _proxies:
        return None
    _proxy_idx = random.randint(0, len(_proxies) - 1)
    line = _proxies[_proxy_idx]
    parts = line.split(":")
    print(f"[PROXY] Using proxy #{_proxy_idx + 1}: {parts[2]}@{parts[0]}:{parts[1]}")
    return line


def _make_session(proxy_line: str | None) -> requests.Session:
    proxy_url = None
    if proxy_line:
        h, p, u, pw = proxy_line.split(":")
        proxy_url = f"http://{u}:{pw}@{h}:{p}"
    sess = requests.Session(impersonate="chrome146", proxies={"http": proxy_url, "https": proxy_url} if proxy_url else None)
    sess.headers.update({
        "Accept": "application/json, text/plain, */*",
        "Content-Type": "application/json",
        "Origin": ORIGIN,
    })
    return sess


def _qvc_headers(token: str | None, referer: str) -> dict:
    h = {
        "X-QVC-Date-Time-Zone": str(int(time.time() * 1000)),
        "X-QVC-Request-Id": f"Default:{uuid.uuid4()}",
        "Referer": referer,
    }
    if token:
        h["Authorization"] = f"Bearer {token}"
    return h


# ── CAPTCHA solving ───────────────────────────────────────────────────────────

_ocr_reader = None
_cnn_model  = None
_CNN_DEVICE = "cpu"
_CNN_CKPT   = os.path.join(os.path.dirname(__file__), "captcha_solver", "captcha_model.pth")
_CNN_TF     = None

if _CNN_AVAILABLE:
    _CNN_DEVICE = "cuda" if _torch.cuda.is_available() else "cpu"
    import torchvision.transforms as _T
    _CNN_TF = _T.Compose([
        _T.Grayscale(), _T.Resize((50, 130)), _T.ToTensor(),
        _T.Normalize((0.5,), (0.5,)),
    ])


def _get_ocr() -> easyocr.Reader:
    global _ocr_reader
    if _ocr_reader is None:
        print("[OCR] Initializing EasyOCR...")
        _ocr_reader = easyocr.Reader(["en"], gpu=False)
    return _ocr_reader


def _get_cnn():
    global _cnn_model
    if _cnn_model is None and _CNN_AVAILABLE and os.path.exists(_CNN_CKPT):
        print("[CNN] Loading captcha model...")
        m = _CaptchaCNN().to(_CNN_DEVICE)
        m.load_state_dict(_torch.load(_CNN_CKPT, map_location=_CNN_DEVICE))
        m.eval()
        _cnn_model = m
    return _cnn_model


def _ocr_solve(img_bytes: bytes) -> str | None:
    try:
        img = Image.open(io.BytesIO(img_bytes)).convert("RGB")
        arr = np.array(img)
        r, g, b = arr[:, :, 0].astype(int), arr[:, :, 1].astype(int), arr[:, :, 2].astype(int)
        saturation = np.maximum(np.maximum(r, g), b) - np.minimum(np.minimum(r, g), b)
        arr[saturation > 40] = [255, 255, 255]
        cleaned = Image.fromarray(arr.astype(np.uint8)).convert("L")
        cleaned = cleaned.resize((cleaned.width * 4, cleaned.height * 4), Image.LANCZOS)
        cleaned = ImageEnhance.Contrast(cleaned).enhance(4.0)
        cleaned = cleaned.filter(ImageFilter.SHARPEN).filter(ImageFilter.SHARPEN)
        buf = io.BytesIO()
        cleaned.save(buf, format="PNG")
        results = _get_ocr().readtext(
            buf.getvalue(), detail=1,
            allowlist="ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789",
            paragraph=False, text_threshold=0.3, low_text=0.3,
        )
        best_text, best_conf = "", 0.0
        for (_, t, c) in results:
            t = t.strip().replace(" ", "")
            if 4 <= len(t) <= 6 and c > best_conf:
                best_text, best_conf = t, c
        if not best_text:
            best_text = "".join(r[1] for r in results).strip().replace(" ", "")
        best_text = best_text[:5]
        print(f"[OCR] '{best_text}' (conf={best_conf:.2f})")
        return best_text or None
    except Exception as e:
        print(f"[OCR] Error: {e}")
        return None


CAPMONSTER_KEY = "757f722f146c06980f5c3b486836beb1"

def solve_captcha(img_bytes: bytes) -> str | None:
    import urllib.request as _urllib
    import json as _json

    b64 = base64.b64encode(img_bytes).decode()

    # Submit task
    create_body = _json.dumps({
        "clientKey": CAPMONSTER_KEY,
        "task": {"type": "ImageToTextTask", "body": b64}
    }).encode()
    req = _urllib.Request(
        "https://api.capmonster.cloud/createTask",
        data=create_body,
        headers={"Content-Type": "application/json"},
        method="POST"
    )
    try:
        with _urllib.urlopen(req, timeout=15) as r:
            resp = _json.loads(r.read())
    except Exception as e:
        print(f"[CAPMONSTER] createTask error: {e}")
        return None

    if resp.get("errorId", 1) != 0:
        print(f"[CAPMONSTER] createTask failed: {resp.get('errorDescription')}")
        return None

    task_id = resp["taskId"]
    print(f"[CAPMONSTER] Task {task_id} submitted, waiting...")

    # Poll for result
    poll_body = _json.dumps({"clientKey": CAPMONSTER_KEY, "taskId": task_id}).encode()
    for _ in range(30):
        time.sleep(1)
        req2 = _urllib.Request(
            "https://api.capmonster.cloud/getTaskResult",
            data=poll_body,
            headers={"Content-Type": "application/json"},
            method="POST"
        )
        try:
            with _urllib.urlopen(req2, timeout=15) as r:
                result = _json.loads(r.read())
        except Exception as e:
            print(f"[CAPMONSTER] poll error: {e}")
            continue

        if result.get("status") == "ready":
            text = result.get("solution", {}).get("text", "").strip()
            print(f"[CAPMONSTER] Solved: '{text}'")
            return text if text else None

    print(f"[CAPMONSTER] Timed out waiting for solution")
    return None


# ── Month date helpers ────────────────────────────────────────────────────────

def _month_range(year: int, month: int) -> tuple[str, str, int, int]:
    last_day = calendar.monthrange(year, month)[1]
    from_utc = datetime(year, month, 1, 0, 0, 0) - PKT_OFFSET
    to_utc   = datetime(year, month, last_day, 0, 0, 0) - PKT_OFFSET
    return (
        from_utc.strftime("%Y-%m-%dT%H:%M:%S.000Z"),
        to_utc.strftime("%Y-%m-%dT%H:%M:%S.000Z"),
        month - 1,
        year,
    )


def _parse_date(s: str) -> datetime:
    return datetime.strptime(s, "%Y-%m-%d")


def _time_strip_ampm(t: str) -> str:
    return t.replace(" AM", "").replace(" PM", "").strip()


# ── API call helpers ──────────────────────────────────────────────────────────

class ApiError(Exception):
    pass

class RateLimitError(Exception):
    pass

class CaptchaError(Exception):
    pass

class SessionExpiredError(Exception):
    pass

class TokenActiveError(Exception):
    pass

# Persists across sessions within this process run
_cached_visa_data: dict | None = None


def _post(sess, path, payload, token, referer, timeout=20):
    url = BASE_URL + path
    headers = _qvc_headers(token, referer)
    try:
        r = sess.post(url, json=payload, headers=headers, timeout=timeout)
        if r.status_code == 429:
            raise RateLimitError("429 Too Many Requests")
        if r.status_code not in (200, 201):
            raise ApiError(f"HTTP {r.status_code} from {path}")
        return r.json()
    except requests.RequestsError as e:
        raise ApiError(f"Connection error on {path}: {e}")


def _get(sess, path, token, referer, params=None, timeout=20):
    url = BASE_URL + path
    headers = _qvc_headers(token, referer)
    try:
        r = sess.get(url, headers=headers, params=params, timeout=timeout)
        if r.status_code == 429:
            raise RateLimitError("429 Too Many Requests")
        if r.status_code not in (200, 201):
            raise ApiError(f"HTTP {r.status_code} from {path}")
        return r.json()
    except requests.RequestsError as e:
        raise ApiError(f"Connection error on {path}: {e}")


# ── QVC API calls ─────────────────────────────────────────────────────────────

def get_token(sess) -> str:
    url = BASE_URL + "/qvc/common/token"
    headers = {
        "Accept": "application/json, text/plain, */*",
        "Authorization": "Bearer null",
        "Content-Type": "application/json",
        "Referer": ORIGIN + "/book",
        "X-QVC-Date-Time-Zone": str(int(time.time() * 1000)),
        "X-QVC-Request-Id": f"Default:{uuid.uuid4()}",
    }
    try:
        r = sess.get(url, headers=headers, params={"countryCode": "undefined"}, timeout=20)
        if r.status_code == 429:
            raise RateLimitError("429 Too Many Requests")
        if r.status_code not in (200, 201):
            raise ApiError(f"HTTP {r.status_code}")
        resp = r.json()
    except requests.RequestsError as e:
        raise ApiError(f"Connection error on token: {e}")
    token = resp.get("token")
    if not token:
        raise ApiError(f"No token in response: {resp}")
    print(f"[TOKEN] Got JWT (expires ~20 min)")
    return token


def get_captcha(sess, token: str) -> dict:
    """Book appointment uses empty body for populateCaptcha (not encrypted)."""
    body = _enc({})
    resp = _post(sess, "/qvc/common/populateCaptcha", body, token, ORIGIN + "/book")
    return _dec(resp)


def validate_visa_and_pass(sess, token: str, captcha_to: dict,
                            captcha_text: str, enc_visa: str, enc_pass: str) -> dict:
    """
    Book appointment validation endpoint.
    Returns visaHolderInfos with full applicant details from Qatar system.
    """
    payload = {
        "passportTO": [
            {
                "countryId": 3050,
                "appointmentType": "Normal",
                "encodevisaNumber": enc_visa,
                "encodepassportNumber": enc_pass,
                "processType": "Bio",
                "scheduleType": "SCHEDULE",
                "languageId": 1,
            }
        ],
        "captchaTO": {
            "statusCode": "OK",
            "passportInfoValidation": None,
            "appRefNo": None,
            "appType": None,
            "noOfApplicant": None,
            "countryId": None,
            "vscId": None,
            "token": token,
            "messageCode": None,
            "visaNumber": None,
            "passportNumber": None,
            "icrResponse": None,
            "captchaId": captcha_to.get("captchaId"),
            "captchaValue": captcha_text,
            "imagePath": None,
            "imageString": captcha_to.get("imageString"),
            "encodevisaNumber": None,
            "encodepassportNumber": None,
            "countryCode": None,
        },
    }
    payload["countryId"] = 3050
    payload["languageId"] = 1
    body = _enc(payload)
    resp = _post(sess, "/qvc/schedule/validatevisaandpass", body, token, ORIGIN + "/book")
    data = _dec(resp)
    print(f"[DEBUG] validatevisaandpass raw: statusCode={data.get('statusCode')!r} messageCode={data.get('messageCode')!r} message={data.get('message')!r} moiErrors={data.get('moiErrorMessage')!r}")

    global _cached_visa_data
    status = data.get("statusCode", "")
    msg = data.get("message", "") or ""

    # E013 = server already has an active session for this visa
    moi_errors = data.get("moiErrorMessage") or []
    for err in moi_errors:
        if err.get("messageCode") == "E013":
            print(f"[E013] Active session on server")
            raise TokenActiveError("E013")

    if "captcha" in msg.lower() or "invalid" in msg.lower():
        raise CaptchaError(f"Invalid captcha: {msg}")
    if status not in ("OK", "200 OK"):
        raise ApiError(f"validatevisaandpass failed: {msg} — {data}")

    pass_info = data.get("passportInfoValidation") or []
    if pass_info and pass_info[0].get("status") != "Y":
        raise ApiError(f"Passport validation failed: {pass_info[0]}")

    visa_infos = data.get("visaHolderInfos") or []
    if not visa_infos:
        raise ApiError(f"No visaHolderInfos returned — wrong credentials? msg={msg!r}")

    print(f"[VALID] Passport validation OK — {len(visa_infos)} applicant(s) found")
    _cached_visa_data = data
    return data


def delete_old_token(sess, token: str, enc_visa: str = None, enc_pass: str = None):
    payload = {"token": "token", "visaNumber": VISA_NUMBER, "passportNumber": PASSPORT_NUMBER}
    body = _enc(payload)
    try:
        raw = _post(sess, "/qvc/schedule/deleteOldToken", body, token, ORIGIN + "/schedule")
        decoded = _dec(raw) if raw else raw
        print(f"[E013] deleteOldToken OK — {decoded.get('statusCode')}")
    except Exception as e:
        print(f"[E013] deleteOldToken failed: {e}")


def token_validation(sess, token: str):
    """Plain JSON POST — no encryption. Called after validatevisaandpass."""
    payload = {
        "visaNumber": VISA_NUMBER,
        "passportNumber": PASSPORT_NUMBER,
        "token": token,
    }
    try:
        _post(sess, "/qvc/schedule/tokenValidation", payload, token, ORIGIN + "/book")
        print("[TOKEN] tokenValidation OK")
    except Exception as e:
        print(f"[TOKEN] tokenValidation skipped: {e}")


def get_vsc_details(sess, token: str, visa_type_id: int) -> dict:
    body = _enc({"languageName": "English", "processType": "Bio",
                 "visaTypeId": visa_type_id, "countryId": 3050})
    resp = _post(sess, "/qvc/schedule/getVscDetails", body, token, ORIGIN + "/book")
    return _dec(resp)


def get_fees(sess, token: str, vsc_id: int):
    payload = {"isAmountPaid": "N", "countryId": 3050, "vscId": vsc_id}
    try:
        _post(sess, "/getfees", payload, token, ORIGIN + "/book")
        print(f"[FEES] getfees OK for vscId={vsc_id}")
    except Exception as e:
        print(f"[FEES] getfees skipped: {e}")


def get_appt_fees(sess, token: str, vsc_id: int):
    payload = {"isAmountPaid": "N", "countryId": 3050, "vscId": vsc_id,
               "apptCatogery": "Normal", "noOfApplicants": 1}
    try:
        _post(sess, "/getapptfees", payload, token, ORIGIN + "/book")
    except Exception as e:
        print(f"[FEES] getapptfees skipped: {e}")


def get_vsc_holidays(sess, token: str, vsc_id: int) -> list:
    body = _enc({"vscId": vsc_id})
    resp = _post(sess, "/qvc/schedule/getVscHoliDays", body, token, ORIGIN + "/book")
    return _dec(resp).get("holidayDates") or []


def get_vsc_weekly_off(sess, token: str, vsc_id: int) -> list:
    body = _enc({"vscId": vsc_id})
    resp = _post(sess, "/qvc/schedule/getVscWeeklyOff", body, token, ORIGIN + "/book")
    return _dec(resp).get("weeklyOffDays") or []


def get_appointment_dates(sess, token: str, year: int, month: int,
                           vsc_id: int, visa_type_id: int,
                           enc_visa: str, enc_pass: str,
                           sponsor_type_ids: list) -> list:
    from_date, to_date, slot_month, slot_year = _month_range(year, month)
    body = _enc({
        "fromDate": from_date,
        "toDate": to_date,
        "slotMonth": slot_month,
        "slotYear": slot_year,
        "bookingDate": "21-MAY-2019",       # placeholder — same as Chrome
        "bookingMode": "Online",
        "countryId": 3050,
        "noOfApplicants": 1,
        "processType": "Bio",
        "visaTypeId": visa_type_id,
        "vscId": vsc_id,
        "appointmentType": "Normal",
        "visaNumber": _DUMMY_VISA,           # dummy — same as Chrome
        "passportNumber": _DUMMY_PASS,
        "encodevisaNumber": enc_visa,
        "encodepassportNumber": enc_pass,
        "sponsorTypeIds": sponsor_type_ids,
    })
    resp = _post(sess, "/qvc/common/getvscappointmentdates", body, token, ORIGIN + "/book")
    data = _dec(resp)
    msg = data.get("message", "")
    if "session expired" in msg.lower():
        raise SessionExpiredError(msg)
    vsc_date_to = data.get("vscAvailableDateTO") or {}
    dates = vsc_date_to.get("availableDates") or []
    max_date = vsc_date_to.get("maxDate", "?")
    print(f"[API] getvscappointmentdates → dates={dates}  maxDate={max_date}")
    return dates


def fetch_slots(sess, token: str, booking_date: str, vsc_id: int,
                enc_visa: str, enc_pass: str, sponsor_type_ids: list) -> list:
    body = _enc({
        "bookingDate": booking_date,
        "bookingMode": "Online",
        "countryId": 3050,
        "noOfApplicants": 1,
        "processType": "Bio",
        "appointmentType": "Normal",
        "vscId": vsc_id,
        "visaNumber": _DUMMY_VISA,           # dummy — same as Chrome
        "passportNumber": _DUMMY_PASS,
        "encodevisaNumber": enc_visa,
        "encodepassportNumber": enc_pass,
        "sponsorTypeIds": sponsor_type_ids,
    })
    resp = _post(sess, "/qvc/common/fetchslot", body, token, ORIGIN + "/book")
    return _dec(resp).get("slotDisplayTOList") or []


def check_slot_available(sess, token: str, schedule_to: dict,
                          slot_quota_id: str, enc_visa: str, enc_pass: str) -> int | None:
    body = _enc({
        "iswaitlistSlot": False,
        "numberOfApplicant": 1,
        "slotQuotaIds": [slot_quota_id],
        "visaNumber": _DUMMY_VISA,           # dummy — same as Chrome
        "passportNumber": _DUMMY_PASS,
        "encodevisaNumber": enc_visa,
        "encodepassportNumber": enc_pass,
        "scheduleTO": schedule_to,
    })
    resp = _post(sess, "/qvc/common/checkslotavailable", body, token, ORIGIN + "/book")
    data = _dec(resp)
    if data.get("isAttemptsExceed") == "Y":
        raise ApiError("Attempts exceeded for checkslotavailable")
    uid = data.get("uniqueIdentifier")
    return int(uid) if uid is not None else None


def get_schedule_captcha(sess, token: str, enc_visa: str, enc_pass: str,
                          captcha_id: str | None = None) -> dict:
    payload: dict = {"visaNumber": enc_visa, "passportNumber": enc_pass}
    if captcha_id:
        payload["captchaId"] = captcha_id
    body = _enc(payload)
    resp = _post(sess, "/qvc/common/populateScheduleCaptcha", body,
                 token, ORIGIN + "/book/reviewsummary")
    return _dec(resp)


def validate_schedule_captcha(sess, token: str, enc_visa: str, enc_pass: str,
                               captcha_id: str, captcha_value: str) -> bool:
    payload = {"visaNumber": enc_visa, "passportNumber": enc_pass,
               "captchaId": captcha_id, "captchaValue": captcha_value}
    body = _enc(payload)
    resp = _post(sess, "/qvc/common/populateScheduleCaptcha", body,
                 token, ORIGIN + "/book/reviewsummary")
    data = _dec(resp)
    status = data.get("captchaStatus", "")
    print(f"[CAPTCHA2] captchaStatus={status!r}")
    return status == "Y"


def save_booking(sess, token: str, booking_payload: dict) -> dict:
    body = _enc(booking_payload)
    resp = _post(sess, "/qvc/schedule/save", body, token, ORIGIN + "/book/reviewsummary")
    return _dec(resp)


# ── Build scheduleTO for book appointment ─────────────────────────────────────

def _build_schedule_to(visa_holder_info: dict, target_vsc: dict,
                        booking_date: str, booking_time: str,
                        example_phone: str = None) -> dict:
    applicant_to = dict(visa_holder_info)
    applicant_to["passportNumber"]    = _DUMMY_PASS
    applicant_to["visaNumber"]        = _DUMMY_VISA
    applicant_to["email"]             = EMAIL
    applicant_to["isPrimaryApplicant"] = "Y"
    applicant_to["isQatarResident"]   = "N"

    application_to = {
        "applicant": [{"isFeePaid": "N", "paidAmount": 0, "isQatarResident": "N"}],
        "bookingMode": "Online",
        "processType": "Bio",
        "numberOfApplicants": 1,
        "isBiometricFeePaid": "N",
        "status": [],
        "isTaxExempted": "N",
        "isCheckVal": True,
        "appointmentStatus": "1",
        "applicantType": "Individual",
        "noshow": False,
        "primaryPhoneNumber": MOBILE_NUMBER,
        "emailId": EMAIL,
        "appType": 1,
        "visaTypeId": visa_holder_info.get("visaTypeId", 51),
        "appointmentDate": booking_date,
        "appointmentType": "Normal",
        "appointmentTime": _time_strip_ampm(booking_time),
    }

    return {
        "toBeCollected": 0,
        "countryTO": COUNTRY_TO,
        "enableGroupCountryTo": {"countryCode": "", "languageName": ""},
        "vscTO": target_vsc,
        "applicationTO": application_to,
        "applicantTOs": [applicant_to],
        "appointmentLetterTo": {"imageList": [{}], "applicantTOs": [], "documentReqs": [{}]},
        "waitlistto": {},
        "feeAmount": 0,
        "feeCurrencyCode": None,
        "normalApptDate": None,
        "examplePhoneNumber": example_phone or MOBILE_NUMBER,
    }


# ── Build save payload ────────────────────────────────────────────────────────

def _build_save_payload(visa_holder_info: dict, target_vsc: dict,
                         booking_date: str, booking_time: str,
                         slot_quota_seq_no: int,
                         captcha_to: dict = None, captcha_value: str = None,
                         example_phone: str = None) -> dict:
    start_time = _time_strip_ampm(booking_time)

    applicant_to = dict(visa_holder_info)
    applicant_to["passportNumber"]     = _DUMMY_PASS
    applicant_to["visaNumber"]         = _DUMMY_VISA
    applicant_to["email"]              = EMAIL
    applicant_to["isPrimaryApplicant"] = "Y"
    applicant_to["isQatarResident"]    = "N"

    country_to = dict(COUNTRY_TO)
    country_to["showLandingPopup"] = "N"

    captcha_payload = None
    if captcha_to:
        captcha_payload = {
            "statusCode": "OK",
            "passportInfoValidation": None,
            "appRefNo": None,
            "appType": None,
            "noOfApplicant": None,
            "countryId": None,
            "vscId": None,
            "token": None,
            "messageCode": None,
            "visaNumber": VISA_NUMBER,
            "passportNumber": PASSPORT_NUMBER,
            "icrResponse": None,
            "captchaId": captcha_to.get("captchaId"),
            "captchaValue": captcha_value or "",
            "imagePath": None,
            "imageString": captcha_to.get("imageString"),
            "encodevisaNumber": None,
            "encodepassportNumber": None,
            "countryCode": None,
        }

    return {
        "toBeCollected": 0,
        "countryTO": country_to,
        "enableGroupCountryTo": {"countryCode": "", "languageName": ""},
        "vscTO": target_vsc,
        "applicationTO": {
            "applicant": [{"isFeePaid": "N", "paidAmount": 0, "isQatarResident": "N"}],
            "bookingMode": "Online",
            "processType": "Bio",
            "numberOfApplicants": 1,
            "isBiometricFeePaid": "Y",
            "status": [],
            "isTaxExempted": "N",
            "isCheckVal": True,
            "appointmentStatus": "1",
            "applicantType": "Individual",
            "appointmentDate": booking_date,
            "appointmentTime": start_time,
            "appointmentType": "Normal",
            "appType": 1,
            "emailId": EMAIL,
            "primaryPhoneNumber": MOBILE_NUMBER,
            "visaTypeId": visa_holder_info.get("visaTypeId", 51),
            "noshow": False,
        },
        "applicantTOs": [applicant_to],
        "appointmentLetterTo": {"imageList": [{}], "applicantTOs": [], "documentReqs": [{}]},
        "waitlistto": {},
        "feeAmount": 0,
        "feeCurrencyCode": None,
        "normalApptDate": None,
        "examplePhoneNumber": example_phone or MOBILE_NUMBER,
        "workingDays": None,
        "workingHours": None,
        "telephone": None,
        "feeTOs": None,
        "slotQuotaSeqNo": slot_quota_seq_no,
        "captchaTO": captcha_payload,
    }


# ── Captcha auto-collection ───────────────────────────────────────────────────

_CAPTCHA_DIR = os.path.join(os.path.dirname(__file__), "captcha_solver", "real_captchas")
_RETRAIN_THRESHOLD = 50
_new_captcha_count = 0


def _save_captcha(img_bytes: bytes, label: str):
    global _new_captcha_count
    try:
        os.makedirs(_CAPTCHA_DIR, exist_ok=True)
        existing = [f for f in os.listdir(_CAPTCHA_DIR) if f.endswith(".png")]
        idx = len(existing) + 1
        label_clean = label.strip().upper()[:5]
        filename = f"{label_clean}_{idx:04d}.png"
        with open(os.path.join(_CAPTCHA_DIR, filename), "wb") as f:
            f.write(img_bytes)
        _new_captcha_count += 1
        print(f"[COLLECT] Saved captcha #{_new_captcha_count}: {filename}  (total={idx})")
        if _new_captcha_count % _RETRAIN_THRESHOLD == 0:
            _retrain_model()
    except Exception as e:
        print(f"[COLLECT] Failed to save captcha: {e}")


def _retrain_model():
    train_script = os.path.join(os.path.dirname(__file__), "captcha_solver", "train.py")
    if not os.path.exists(train_script):
        return
    import subprocess
    python = os.path.join(os.path.dirname(__file__), "venv312", "Scripts", "python.exe")
    print(f"[RETRAIN] Starting retraining ({_new_captcha_count} new captchas)...")
    subprocess.Popen([python, train_script], cwd=os.path.dirname(__file__))


# ── Discord notifications ─────────────────────────────────────────────────────

def _notify_discord(message: str, webhook: str = None):
    url = webhook or DISCORD_WEBHOOK
    if not url:
        return
    try:
        requests.post(url, json={"content": message}, timeout=10)
    except Exception as e:
        print(f"[DISCORD] Failed to send notification: {e}")


_ALERT_COOLDOWN = 30 * 60  # 30 minutes
_alerted_slots: dict = {}  # (date_str, time_str) -> last alert timestamp
_alerted_no_slots: dict = {}
_NO_SLOT_COOLDOWN = 30 * 60  # 30 minutes


def _should_alert(center: str, date_str: str) -> bool:
    key = (center, date_str)
    last = _alerted_slots.get(key)
    if last is None or time.time() - last >= _ALERT_COOLDOWN:
        _alerted_slots[key] = time.time()
        return True
    return False


def _should_alert_no_slot(center: str, date_str: str) -> bool:
    key = (center, date_str)
    last = _alerted_no_slots.get(key)
    if last is None or time.time() - last >= _NO_SLOT_COOLDOWN:
        _alerted_no_slots[key] = time.time()
        return True
    return False


# ── Date filtering ────────────────────────────────────────────────────────────

def _is_before_urgent(date_str: str) -> bool:
    """Return True if date_str is strictly before URGENT_MEDICAL_DATE (or no urgent date set)."""
    if not URGENT_MEDICAL_DATE:
        return True
    try:
        return datetime.strptime(date_str, "%Y-%m-%d") < datetime.strptime(URGENT_MEDICAL_DATE, "%Y-%m-%d")
    except ValueError:
        return True


def _get_year_months() -> list[tuple[int, int]]:
    now = datetime.now()
    result = []
    for name in MONTHS_TO_CHECK:
        m = MONTH_NAMES.get(name)
        if m is None:
            continue
        y = now.year
        if m < now.month:
            y += 1
        result.append((y, m))
    return result


def _dates_worker(center_name, vsc_id, year, month, token, proxy_line,
                  visa_type_id, enc_visa, enc_pass, sponsor_type_ids):
    t_sess = _make_session(proxy_line)
    return get_appointment_dates(
        t_sess, token, year, month, vsc_id,
        visa_type_id, enc_visa, enc_pass, sponsor_type_ids
    )


def _slots_worker(date_str, vsc_id, token, proxy_line, enc_visa, enc_pass, sponsor_type_ids):
    t_sess = _make_session(proxy_line)
    return fetch_slots(t_sess, token, date_str, vsc_id, enc_visa, enc_pass, sponsor_type_ids)


def run():
    _notify_discord(
        f"\U0001f7e2 **QVC Bot Started**\n"
        f"\U0001f4cd Centers: **Islamabad, Karachi**\n"
        f"\U0001f4c5 Scanning: {', '.join(MONTHS_TO_CHECK)}",
        webhook=DISCORD_SERVER_WEBHOOK,
    )
    _load_proxies()

    enc_visa  = _encode_url_safe(VISA_NUMBER)
    enc_pass  = _encode_url_safe(PASSPORT_NUMBER)

    target_vsc = VSC_MAP.get(QVC_CENTER)
    if target_vsc is None:
        print(f"[ERROR] Unknown QVC_CENTER: {QVC_CENTER!r}. Use 'Islamabad' or 'Karachi'.")
        return

    year_months = _get_year_months()
    print(f"[CFG] Center={QVC_CENTER}  Months={MONTHS_TO_CHECK}  "
          f"UrgentBefore={URGENT_MEDICAL_DATE or 'any'}")

    session_num = 0
    while True:
        session_num += 1
        print(f"\n{'='*60}")
        print(f"[SESSION {session_num}] {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} — getting token + validating...")

        proxy_line = _pick_proxy()
        sess = _make_session(proxy_line)

        try:
            # ── Step 1: Token ─────────────────────────────────────────────────
            token = get_token(sess)

            # ── Steps 2+3: Get captcha + solve + validatevisaandpass ──────────
            valid_data = None
            # Clear any pre-existing session once before we start
            delete_old_token(sess, token)
            for cap_attempt in range(12):
                captcha_to = get_captcha(sess, token)
                img_b64 = captcha_to.get("imageString", "")
                if not img_b64:
                    raise ApiError("populateCaptcha returned no imageString")
                img_bytes = base64.b64decode(img_b64)
                print(f"[CAPTCHA1] Solving captcha (attempt {cap_attempt+1})...")
                captcha_text = solve_captcha(img_bytes)
                if not captcha_text or len(captcha_text) < 4:
                    print(f"[CAPTCHA1] OCR failed, retrying...")
                    continue
                print(f"[CAPTCHA1] Answer: '{captcha_text}'")
                try:
                    valid_data = validate_visa_and_pass(
                        sess, token, captcha_to, captcha_text, enc_visa, enc_pass)
                    _save_captcha(img_bytes, captcha_text)
                    break
                except TokenActiveError:
                    print(f"[E013] Session created by wrong captcha — clearing...")
                    delete_old_token(sess, token)
                    continue
                except CaptchaError as e:
                    print(f"[CAPTCHA1] {e} — getting new captcha...")
                    continue

            if valid_data is None:
                raise ApiError("Failed to pass captcha+validation after 8 attempts")

            # ── Step 4: tokenValidation ───────────────────────────────────────
            token_validation(sess, token)

            # Extract applicant info from visaHolderInfos
            visa_infos = valid_data.get("visaHolderInfos") or []
            visa_holder = visa_infos[0] if visa_infos else {}
            visa_type_id = visa_holder.get("visaTypeId", 51)
            sponsor_type_ids = [str(visa_holder.get("sponsorTypeCode", "2"))]
            example_phone = valid_data.get("examplePhoneNumber") or MOBILE_NUMBER

            print(f"[VALID] visaTypeId={visa_type_id}  sponsorTypes={sponsor_type_ids}")

            # ── Step 5: VSC details + fees (required to init server session state) ──
            try:
                get_vsc_details(sess, token, visa_type_id)
                print("[API] getVscDetails OK")
            except Exception as e:
                print(f"[API] getVscDetails skipped: {e}")
            get_fees(sess, token, target_vsc["vscId"])

            # ── Poll loop — runs until server signals session expired ─────────
            poll_num = 0

            while True:
                poll_num += 1
                print(f"\n[POLL {poll_num}] {datetime.now().strftime('%H:%M:%S')}")

                # Phase 1: scan ALL centers + all months in parallel
                all_center_slots = {}
                got_429 = False

                # Step 1a: fetch available dates for all center+month combos in parallel
                _date_scan_tasks = [
                    (cname, cvsc, yr, mo)
                    for cname, cvsc in VSC_MAP.items()
                    for (yr, mo) in year_months
                ]
                _dates_map = {}

                with ThreadPoolExecutor(max_workers=len(_date_scan_tasks)) as _ex:
                    _fmap = {
                        _ex.submit(
                            _dates_worker,
                            cname, cvsc["vscId"], yr, mo, token, proxy_line,
                            visa_type_id, enc_visa, enc_pass, sponsor_type_ids
                        ): (cname, yr, mo)
                        for (cname, cvsc, yr, mo) in _date_scan_tasks
                    }
                    for _fut in as_completed(_fmap):
                        _key = _fmap[_fut]
                        try:
                            _dates_map[_key] = _fut.result() or []
                        except SessionExpiredError as e:
                            raise SessionExpiredError(e)
                        except RateLimitError:
                            print(f"[429] Rate limited on dates scan — switching proxy...")
                            proxy_line = _pick_proxy()
                            sess = _make_session(proxy_line)
                            got_429 = True

                if got_429:
                    continue

                # Step 1b: fetch slots for all (center, date) combos in parallel
                _slot_scan_tasks = [
                    (cname, VSC_MAP[cname], date_str)
                    for (cname, yr, mo), dates in _dates_map.items()
                    for date_str in dates
                ]
                _slots_map = {}

                if _slot_scan_tasks:
                    with ThreadPoolExecutor(max_workers=min(len(_slot_scan_tasks), 8)) as _ex:
                        _fmap = {
                            _ex.submit(
                                _slots_worker,
                                date_str, cvsc["vscId"],
                                token, proxy_line, enc_visa, enc_pass, sponsor_type_ids
                            ): (cname, date_str)
                            for (cname, cvsc, date_str) in _slot_scan_tasks
                        }
                        for _fut in as_completed(_fmap):
                            _key = _fmap[_fut]
                            try:
                                _slots_map[_key] = _fut.result() or []
                            except RateLimitError:
                                print(f"[429] Rate limited on fetchslot — switching proxy...")
                                proxy_line = _pick_proxy()
                                sess = _make_session(proxy_line)
                                got_429 = True

                if got_429:
                    continue

                # Step 1c: build all_center_slots, print results, alert on no-slot dates
                for center_name, center_vsc in VSC_MAP.items():
                    center_urgent = []
                    center_normal = []
                    for (yr, mo) in year_months:
                        dates = _dates_map.get((center_name, yr, mo), [])
                        if not dates:
                            print(f"[SCAN] No dates available for {calendar.month_name[mo]} {yr} at {center_name}")
                            continue
                        print(f"[SCAN] Dates found for {calendar.month_name[mo]} {yr} at {center_name}: {', '.join(dates)}")
                        for date_str in dates:
                            _is_urgent = _is_before_urgent(date_str)
                            _label = "URGENT" if (URGENT_MEDICAL_DATE and _is_urgent) else ("NORMAL" if URGENT_MEDICAL_DATE else "")
                            _tag   = f"[{_label}] " if _label else ""
                            slots  = _slots_map.get((center_name, date_str), [])
                            _dt    = datetime.strptime(date_str, "%Y-%m-%d")
                            _month_name   = _dt.strftime("%B")
                            _date_display = f"{_dt.month}-{_dt.day}-{_dt.year}"
                            if not slots:
                                print(f"{_tag}[{center_name}] Month: {_month_name}  Date: {_date_display}  Time Slots: \U0001f534 Not Available")
                                if _should_alert_no_slot(center_name, date_str):
                                    _notify_discord(
                                        f"\U0001f514 Slots \U0001f7e2 Open\n"
                                        f"\U0001f4cd Center: {center_name}\n"
                                        f"\U0001f4c5 Date : {_month_name}-{_dt.day}-{_dt.year}\n"
                                        f"⏰ No slot available",
                                        webhook=DISCORD_URGENT_WEBHOOK if (URGENT_MEDICAL_DATE and _is_urgent) else DISCORD_WEBHOOK,
                                    )
                                continue
                            _avail_times = []
                            for slot_entry in slots:
                                slot_tos = slot_entry.get("slotTO", [])
                                if not slot_tos:
                                    continue
                                slot_to = slot_tos[0]
                                slot_id = slot_to.get("slotQuotaId")
                                avail   = slot_to.get("available", 0)
                                stime   = slot_entry.get("slotDisplayStartTime", "")
                                etime   = slot_entry.get("slotDisplayEndTime", "")
                                if avail and slot_id:
                                    if _is_urgent:
                                        center_urgent.append((date_str, stime, etime, slot_id, avail))
                                    else:
                                        center_normal.append((date_str, stime, etime, slot_id, avail))
                                    _avail_times.append(stime.replace(" ", ""))
                            _slot_status = ("\U0001f7e2 Open: " + "  ".join(_avail_times)) if _avail_times else "\U0001f534 Not Available"
                            print(f"{_tag}\U0001f4cd {center_name}  Month: {_month_name}  Date: {_date_display}  Time Slots: {_slot_status}")
                            if not _avail_times:
                                if _should_alert_no_slot(center_name, date_str):
                                    _notify_discord(
                                        f"\U0001f514 Slots \U0001f7e2 Open\n"
                                        f"\U0001f4cd Center: {center_name}\n"
                                        f"\U0001f4c5 Date : {_month_name}-{_dt.day}-{_dt.year}\n"
                                        f"⏰ No slot available",
                                        webhook=DISCORD_URGENT_WEBHOOK if (URGENT_MEDICAL_DATE and _is_urgent) else DISCORD_WEBHOOK,
                                    )
                    all_center_slots[center_name] = {"urgent": center_urgent, "normal": center_normal}

                # Phase 2: Discord alerts per center
                any_slots_found = False
                _urgent_parts = []
                _normal_parts = []
                for center_name, slots_dict in all_center_slots.items():
                    c_urgent = slots_dict["urgent"]
                    c_normal = slots_dict["normal"]
                    if not c_urgent and not c_normal:
                        continue
                    any_slots_found = True
                    c_urgent.sort(key=lambda x: (x[0], x[1]))
                    c_normal.sort(key=lambda x: (x[0], x[1]))

                    print(f"\n{'─'*50}")
                    if c_urgent:
                        print(f"[URGENT] {len(c_urgent)} slot(s) at {center_name}:")
                        for (ds, st, et, sid, av) in c_urgent:
                            _dtu = datetime.strptime(ds, "%Y-%m-%d")
                            print(f"[URGENT] [{center_name}] Month: {_dtu.strftime('%B')}  Date: {_dtu.month}-{_dtu.day}-{_dtu.year}  Time: {st.replace(' ','')}")
                    if c_normal:
                        print(f"[NORMAL] {len(c_normal)} slot(s) at {center_name}:")
                        for (ds, st, et, sid, av) in c_normal:
                            _dtn = datetime.strptime(ds, "%Y-%m-%d")
                            print(f"[NORMAL] [{center_name}] Month: {_dtn.strftime('%B')}  Date: {_dtn.month}-{_dtn.day}-{_dtn.year}  Time: {st.replace(' ','')}")
                    print(f"{'─'*50}")

                    _urgent_dates = {ds for (ds, *_) in c_urgent if _should_alert(center_name, ds)}
                    _normal_dates = {ds for (ds, *_) in c_normal if _should_alert(center_name, ds)}
                    _new_urgent = [(ds, st, et, sid, av) for (ds, st, et, sid, av) in c_urgent if ds in _urgent_dates]
                    _new_normal = [(ds, st, et, sid, av) for (ds, st, et, sid, av) in c_normal if ds in _normal_dates]

                    if _new_urgent:
                        _u_d = {}
                        for (ds, st, et, sid, av) in _new_urgent:
                            _u_d.setdefault(ds, []).append(st.replace(" ", ""))
                        _upart = f"\U0001f514 **URGENT SLOTS AVAILABLE**\n"
                        _upart += f"\U0001f4cd Center: {center_name}\n"
                        _upart += "\U0001f534 Before: " + URGENT_MEDICAL_DATE + "\n"
                        for ds, times in _u_d.items():
                            _udt = datetime.strptime(ds, "%Y-%m-%d")
                            _upart += f"\U0001f4c5 Date : {_udt.strftime('%B')}-{_udt.day}-{_udt.year}\n"
                            _upart += ("⏰  " + "   ".join(times) + "\n") if times else "⏰ No Time Slot available\n"
                        _urgent_parts.append(_upart.strip())

                    if _new_normal:
                        _n_d = {}
                        for (ds, st, et, sid, av) in _new_normal:
                            _n_d.setdefault(ds, []).append(st.replace(" ", ""))
                        _npart = f"\U0001f514 Slots \U0001f7e2 Open\n"
                        _npart += f"\U0001f4cd Center: {center_name}\n"
                        for ds, times in _n_d.items():
                            _ndt = datetime.strptime(ds, "%Y-%m-%d")
                            _npart += f"\U0001f4c5 Date : {_ndt.strftime('%B')}-{_ndt.day}-{_ndt.year}\n"
                            _npart += ("⏰  " + "   ".join(times) + "\n") if times else "⏰ No Time Slot available\n"
                        _normal_parts.append(_npart.strip())

                if _urgent_parts:
                    _notify_discord("@everyone\n\n" + "\n\n".join(_urgent_parts), webhook=DISCORD_URGENT_WEBHOOK)
                if _normal_parts:
                    _notify_discord("@everyone\n\n" + "\n\n".join(_normal_parts), webhook=DISCORD_WEBHOOK)

                if not any_slots_found:
                    print(f"[POLL] No slots found. Next scan in {POLL_INTERVAL}s...")
                    time.sleep(POLL_INTERVAL)
                    continue

                # Phase 3 setup: extract booking slots from target center only
                urgent_slots = all_center_slots.get(QVC_CENTER, {}).get("urgent", [])
                normal_slots = all_center_slots.get(QVC_CENTER, {}).get("normal", [])
                all_slots = urgent_slots + normal_slots
                slots_to_book = urgent_slots if URGENT_MEDICAL_DATE else all_slots
                if not slots_to_book:
                    print(f"[POLL] No URGENT slots found before {URGENT_MEDICAL_DATE}. Waiting {POLL_INTERVAL}s...")
                    time.sleep(POLL_INTERVAL)
                    continue

                _eb = slots_to_book[0]
                print(f"[BOOKING] Attempting earliest: {_eb[0]}  {_eb[1]}")

                # Phase 3: book earliest slot
                found = False
                for (date_str, start_time, end_time, slot_id, _avail) in slots_to_book:
                    booking_time = _time_strip_ampm(start_time)
                    schedule_to  = _build_schedule_to(visa_holder, target_vsc, date_str, booking_time, example_phone)

                    try:
                        uid = check_slot_available(sess, token, schedule_to, slot_id, enc_visa, enc_pass)
                    except RateLimitError:
                        print(f"[429] Rate limited on checkslot — switching proxy...")
                        proxy_line = _pick_proxy()
                        sess = _make_session(proxy_line)
                        continue
                    except ApiError as e:
                        print(f"[SLOT] checkslotavailable error: {e}")
                        continue
                    if uid is None:
                        print(f"[SLOT] {date_str} {start_time} — no longer available, trying next...")
                        continue
                    print(f"[SLOT] Confirmed: {date_str} {start_time}  uid={uid}")

                    # Captcha 2 — 5 attempts per cycle; alert + 10s wait + fresh captcha on failure
                    print("[CAPTCHA2] Getting schedule captcha...")
                    sch_cap_to = None
                    cap2_valid = False
                    cap2_text = ""

                    for cap2_cycle in range(5):
                        for cap2_attempt in range(5):
                            sch_cap_to = get_schedule_captcha(
                                sess, token, enc_visa, enc_pass,
                                sch_cap_to.get("captchaId") if sch_cap_to else None,
                            )
                            img2_b64 = sch_cap_to.get("imageString", "")
                            if not img2_b64:
                                break
                            cap2_text = solve_captcha(base64.b64decode(img2_b64))
                            if not cap2_text:
                                print(f"[CAPTCHA2] Could not solve (cycle {cap2_cycle+1}, attempt {cap2_attempt+1}/5), refreshing...")
                                continue
                            print(f"[CAPTCHA2] Answer: '{cap2_text}' (cycle {cap2_cycle+1}, attempt {cap2_attempt+1}/5)")
                            try:
                                cap2_valid = validate_schedule_captcha(
                                    sess, token, enc_visa, enc_pass,
                                    sch_cap_to["captchaId"], cap2_text)
                            except ApiError as e:
                                print(f"[CAPTCHA2] Validation error: {e}")
                            if cap2_valid:
                                break
                            print(f"[CAPTCHA2] Wrong answer...")

                        if cap2_valid:
                            break

                        _total = (cap2_cycle + 1) * 5
                        print(f"[CAPTCHA2] {_total} attempts failed — alerting, waiting 10s for mobile...")
                        _notify_discord(
                            f"⚠️ **Captcha 2 Failed {_total} Times**\n"
                            f"📅 Slot: {date_str}  ⏰ {start_time}\n"
                            f"Open site on mobile and select center within 10s...",
                            webhook=DISCORD_SERVER_WEBHOOK,
                        )
                        time.sleep(10)
                        sch_cap_to = None

                    if not cap2_valid:
                        print(f"[CAPTCHA2] All cycles exhausted — still attempting save")

                    save_payload = _build_save_payload(
                        visa_holder, target_vsc, date_str, booking_time, uid,
                        captcha_to=sch_cap_to, captcha_value=cap2_text,
                        example_phone=example_phone,
                    )

                    if DRY_RUN:
                        print(f"[DRY_RUN] Would save: {date_str} {booking_time} at {QVC_CENTER}  uid={uid}")
                        print(f"[DRY_RUN] Payload: {json.dumps(save_payload, indent=2)}")
                        found = True
                        break

                    print(f"[SAVE] Saving booking for {date_str} {booking_time} at {QVC_CENTER}...")
                    try:
                        result = save_booking(sess, token, save_payload)
                    except ApiError as e:
                        print(f"[SAVE] Error: {e}")
                        _notify_discord(
                            f"❌ **Save API Error**\n`{e}`\n📅 {date_str}  ⏰ {booking_time}",
                            webhook=DISCORD_SERVER_WEBHOOK,
                        )
                        continue

                    msg_code = result.get("messageCode", "")
                    status   = result.get("statusCode", "")
                    icr      = result.get("icrResponse") or {}

                    if msg_code == "E016":
                        print("[SAVE] E016 captcha wrong — retrying next slot")
                        _notify_discord(
                            f"⚠️ **Save Error E016** — Captcha wrong\n📅 {date_str}  ⏰ {booking_time}\nRetrying next slot...",
                            webhook=DISCORD_SERVER_WEBHOOK,
                        )
                        continue
                    if msg_code:
                        print(f"[SAVE] Server error: {msg_code}  status={status}")
                        _notify_discord(
                            f"❌ **Save Server Error**\nCode: `{msg_code}`  Status: `{status}`\n📅 {date_str}  ⏰ {booking_time}",
                            webhook=DISCORD_SERVER_WEBHOOK,
                        )
                        continue

                    if status in ("OK", "200 OK") or icr.get("paymentStatus") == "SUCCESS":
                        print("\n" + "★" * 60)
                        print(f"✓  BOOKED!  {date_str} at {booking_time}")
                        print(f"   Center: {QVC_CENTER}")
                        print(f"   Visa:   {VISA_NUMBER}  Pass: {PASSPORT_NUMBER}")
                        print("★" * 60 + "\n")
                        _booked_dt = datetime.strptime(date_str, "%Y-%m-%d")
                        _success_msg = (
                            "\u2705 **APPOINTMENT BOOKED!**\n"
                            f"\U0001f4cd  {QVC_CENTER}\n\n"
                            f"\U0001f4c5  **{_booked_dt.strftime('%B')}**  |  {_booked_dt.month}-{_booked_dt.day}-{_booked_dt.year}\n"
                            f"\u23f0  {booking_time}\n"
                            f"\U0001f6c2  Passport: {PASSPORT_NUMBER}"
                        )
                        _notify_discord(_success_msg)
                        _notify_discord(_success_msg, webhook=DISCORD_SERVER_WEBHOOK)
                        found = True
                        break

                if found:
                    print("[DONE] Appointment booked successfully. Exiting.")
                    return


        except SessionExpiredError:
            print(f"[SESSION] Restarting session immediately...")
        except RateLimitError as e:
            print(f"[429] {e} — switching proxy and retrying in 5s...")
            proxy_line = _pick_proxy()
            sess = _make_session(proxy_line)
            time.sleep(5)
        except (ApiError, CaptchaError) as e:
            err_str = str(e).lower()
            if "connection" in err_str or "timeout" in err_str or "timed out" in err_str or "http 403" in err_str:
                print(f"[ERR] {e} — switching proxy and retrying in 3s...")
                proxy_line = _pick_proxy()
                sess = _make_session(proxy_line)
                time.sleep(3)
            else:
                print(f"[ERR] {e} — retrying in 10s...")
                time.sleep(10)
        except KeyboardInterrupt:
            print("\n[STOP] Interrupted by user.")
            return
        except Exception as e:
            print(f"[ERR] Unexpected: {e} — retrying in 15s...")
            time.sleep(15)


if __name__ == "__main__":
    run()
