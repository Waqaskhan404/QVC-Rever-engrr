"""
QVC Direct API — Reschedule Appointment
=========================================
Calls Qatar Visa Center APIs directly (no browser) to reschedule an appointment.

Usage:
    venv312/Scripts/python qvc_direct_api.py

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
import winsound
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

PASSPORT_NUMBER   = "XC4115103"
VISA_NUMBER       = "382026075537"
MOBILE_NUMBER     = "00923045454166"
EMAIL             = "waqas.khan.40004@gmail.com"

# Which VSC to reschedule TO
QVC_CENTER        = "Islamabad"        # "Islamabad" or "Karachi"

# Months to scan for available slots (order matters)
MONTHS_TO_CHECK   = ["April", "May"]

# Only book a date STRICTLY BEFORE this date. Set to "" or None to book any date.
URGENT_MEDICAL_DATE = ""

# How many seconds to wait between polling cycles (when no slots found)
POLL_INTERVAL     = 10

# DRY_RUN = True → go through the full flow but skip the final save (for testing)
DRY_RUN = False

# Proxy file — one proxy per line: host:port:user:pass
PROXY_FILE = os.path.join(os.path.dirname(__file__), "Webshare residential proxies.txt")

# ── CONSTANTS ─────────────────────────────────────────────────────────────────

BASE_URL  = "https://agent.qatarvisacenter.com"
ORIGIN    = "https://www.qatarvisacenter.com"

DISCORD_WEBHOOK        = "https://discordapp.com/api/webhooks/1495888024561516777/Hdcv7CY-fE8zjtxo3eYupCVLYzapg_cIlY3lbF0YSLWWyor1TIq7hBWYMCn4RsF2TOGO"
DISCORD_URGENT_WEBHOOK = "https://discord.com/api/webhooks/1496147361569833140/-LydSfbfDWM0KWlEjhePMABhEMzgTbp8i0wm_aiHPJ7565KdIMQMFKMZhgQMd7zZnK0Q"

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

AES_KEY = b"cvq@4202temoib!&"
PKT_OFFSET = timedelta(hours=5)

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
    """AES-encrypt a visa/passport number for URL-safe embedding in request fields."""
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
    for _ in range(20):
        time.sleep(2)
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
    """(fromDate, toDate, slotMonth 0-indexed, slotYear) for getvscappointmentdates."""
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
    """'11:45 AM' → '11:45', '14:00 PM' → '14:00'."""
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
        "Referer": ORIGIN + "/manage",
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
    """GET initial reschedule captcha. Returns captchaTO dict with imageString."""
    body = _enc({"token": token})
    resp = _post(sess, "/qvc/common/populateCaptcha", body, token,
                 ORIGIN + "/manage")
    return _dec(resp)


def reschedule_validate_captcha(sess, token: str, captcha_to: dict,
                                 captcha_text: str,
                                 enc_visa: str, enc_pass: str) -> dict:
    """Validate passport/visa + captcha → returns reschedulingTO dict."""
    payload = {
        "statusCode": "OK",
        "token": token,
        "visaNumber": enc_visa,
        "passportNumber": enc_pass,
        "captchaId": captcha_to.get("captchaId"),
        "captchaValue": captcha_text,
        "imagePath": captcha_to.get("imagePath"),
        "imageString": captcha_to.get("imageString"),
        "encodevisaNumber": None,
        "encodepassportNumber": None,
        "countryCode": "PK",
    }
    body = _enc(payload)
    resp = _post(sess, "/qvc/common/rescheduleDetailValidateCaptcha", body,
                 token, ORIGIN + "/manage/reschedule")
    data = _dec(resp)

    # E013 = active session already on server
    moi_errors = data.get("moiErrorMessage") or []
    for err in moi_errors:
        if err.get("messageCode") == "E013":
            print(f"[E013] Active session on server")
            raise TokenActiveError("E013")

    if data.get("statusCode") not in ("OK", "200 OK"):
        msg = data.get("message", "")
        if "captcha" in msg.lower() or "invalid" in msg.lower():
            raise CaptchaError(f"Invalid captcha: {msg}")
        raise ApiError(f"Captcha validation failed: {data.get('messageCode')} — {data}")
    resched = data.get("reschedulingTO") or {}
    if not resched.get("applId"):
        msg = data.get("message", "")
        raise ApiError(f"reschedulingTO empty — wrong credentials or no existing appointment? server msg: {msg!r}")
    return resched


def delete_old_token(sess, token: str, enc_visa: str = None, enc_pass: str = None):
    payload = {"token": "token", "visaNumber": VISA_NUMBER, "passportNumber": PASSPORT_NUMBER}
    body = _enc(payload)
    try:
        raw = _post(sess, "/qvc/schedule/deleteOldToken", body, token, ORIGIN + "/schedule")
        decoded = _dec(raw) if raw else raw
        print(f"[E013] deleteOldToken OK — {decoded.get('statusCode')}")
    except Exception as e:
        print(f"[E013] deleteOldToken failed: {e}")


def get_vsc_details(sess, token: str, country_id: int, visa_type_id: int) -> dict:
    body = _enc({"languageName": "English", "processType": "Bio",
                 "visaTypeId": visa_type_id, "countryId": country_id})
    resp = _post(sess, "/qvc/schedule/getVscDetails", body,
                 token, ORIGIN + "/manage/reschedule")
    return _dec(resp)


def get_vsc_holidays(sess, token: str, vsc_id: int) -> list:
    body = _enc({"vscId": vsc_id})
    resp = _post(sess, "/qvc/schedule/getVscHoliDays", body,
                 token, ORIGIN + "/manage/reschedule")
    data = _dec(resp)
    return data.get("holidayDates") or []


def get_vsc_weekly_off(sess, token: str, vsc_id: int) -> list:
    body = _enc({"vscId": vsc_id})
    resp = _post(sess, "/qvc/schedule/getVscWeeklyOff", body,
                 token, ORIGIN + "/manage/reschedule")
    data = _dec(resp)
    return data.get("weeklyOffDays") or []


def get_appointment_dates(sess, token: str, year: int, month: int,
                           vsc_id: int, country_id: int, visa_type_id: int,
                           enc_visa: str, enc_pass: str,
                           sponsor_type_ids: list) -> list:
    """Returns list of available date strings for the given month."""
    from_date, to_date, slot_month, slot_year = _month_range(year, month)
    body = _enc({
        "fromDate": from_date,
        "toDate": to_date,
        "slotMonth": slot_month,
        "slotYear": slot_year,
        "bookingMode": "Online",
        "countryId": country_id,
        "noOfApplicants": 1,
        "processType": "Bio",
        "visaTypeId": visa_type_id,
        "vscId": vsc_id,
        "appointmentType": "Normal",
        "encodevisaNumber": enc_visa,
        "encodepassportNumber": enc_pass,
        "sponsorTypeIds": sponsor_type_ids,
    })
    resp = _post(sess, "/qvc/common/getvscappointmentdates", body,
                 token, ORIGIN + "/manage/reschedule")
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
                country_id: int, visa_number: str, passport_number: str,
                enc_visa: str, enc_pass: str, sponsor_type_ids: list) -> list:
    """Returns slotDisplayTOList."""
    body = _enc({
        "bookingDate": booking_date,
        "bookingMode": "Online",
        "countryId": country_id,
        "noOfApplicants": 1,
        "processType": "Bio",
        "appointmentType": "Normal",
        "vscId": vsc_id,
        "visaNumber": visa_number,
        "passportNumber": passport_number,
        "encodevisaNumber": enc_visa,
        "encodepassportNumber": enc_pass,
        "sponsorTypeIds": sponsor_type_ids,
    })
    resp = _post(sess, "/qvc/common/fetchslot", body,
                 token, ORIGIN + "/manage/reschedule")
    data = _dec(resp)
    return data.get("slotDisplayTOList") or []


def check_slot_available(sess, token: str, schedule_to: dict,
                          slot_quota_id: str, enc_visa: str,
                          enc_pass: str) -> int | None:
    """Returns uniqueIdentifier (int) or None if slot not available."""
    body = _enc({
        "iswaitlistSlot": False,
        "numberOfApplicant": 1,
        "slotQuotaIds": [slot_quota_id],
        "encodevisaNumber": enc_visa,
        "encodepassportNumber": enc_pass,
        "scheduleTO": schedule_to,
    })
    resp = _post(sess, "/qvc/common/checkslotavailable", body,
                 token, ORIGIN + "/manage/reschedule")
    data = _dec(resp)
    if data.get("isAttemptsExceed") == "Y":
        raise ApiError("Attempts exceeded for checkslotavailable")
    uid = data.get("uniqueIdentifier")
    return int(uid) if uid is not None else None


def get_schedule_captcha(sess, token: str, enc_visa: str, enc_pass: str,
                          captcha_id: str | None = None) -> dict:
    """GET a new schedule captcha image. Returns captchaTO dict."""
    payload: dict = {"visaNumber": enc_visa, "passportNumber": enc_pass}
    if captcha_id:
        payload["captchaId"] = captcha_id
    body = _enc(payload)
    resp = _post(sess, "/qvc/common/populateScheduleCaptcha", body,
                 token, ORIGIN + "/manage/reschedule/reviewsummary")
    return _dec(resp)


def validate_schedule_captcha(sess, token: str, enc_visa: str, enc_pass: str,
                               captcha_id: str, captcha_value: str) -> bool:
    """Validate schedule captcha. Returns True if correct."""
    payload = {
        "visaNumber": enc_visa,
        "passportNumber": enc_pass,
        "captchaId": captcha_id,
        "captchaValue": captcha_value,
    }
    body = _enc(payload)
    resp = _post(sess, "/qvc/common/populateScheduleCaptcha", body,
                 token, ORIGIN + "/manage/reschedule/reviewsummary")
    data = _dec(resp)
    status = data.get("captchaStatus", "")
    print(f"[CAPTCHA2] captchaStatus={status!r}")
    return status == "Y"


def save_reschedule(sess, token: str, reschedule_payload: dict) -> dict:
    body = _enc(reschedule_payload)
    resp = _post(sess, "/qvc/reschedule/save", body,
                 token, ORIGIN + "/manage/reschedule/reviewsummary")
    return _dec(resp)


# ── Build scheduleTO for checkslotavailable ───────────────────────────────────

def _build_schedule_to(resched_to: dict, target_vsc: dict,
                        booking_date: str, booking_time: str) -> dict:
    """Build the scheduleTO from reschedulingTO + chosen date/time/vsc."""
    app_to = dict(resched_to)
    app_to["appointmentDate"] = booking_date
    app_to["appointmentTime"] = _time_strip_ampm(booking_time)
    applicants = resched_to.get("applicant") or []
    # normalApptDate = original appointment date from status[2]
    status_arr = resched_to.get("status") or []
    normal_appt_date = status_arr[2] if len(status_arr) > 2 else datetime.now().strftime("%Y-%m-%d")
    return {
        "toBeCollected": 0,
        "countryTO": COUNTRY_TO,
        "enableGroupCountryTo": {"countryCode": "", "languageName": ""},
        "vscTO": target_vsc,
        "applicationTO": app_to,
        "applicantTOs": applicants,
        "appointmentLetterTo": {"imageList": [{}], "applicantTOs": [], "documentReqs": [{}]},
        "waitlistto": {},
        "feeAmount": 0,
        "feeCurrencyCode": "",
        "normalApptDate": normal_appt_date,
        "feeTOs": resched_to.get("feeTOs") or [],
    }


# ── Build saveReschedule payload ──────────────────────────────────────────────

def _build_save_payload(resched_to: dict, target_vsc: dict,
                         booking_date: str, booking_time: str,
                         slot_quota_seq_no: int,
                         captcha_to: dict = None, captcha_value: str = None) -> dict:
    start_time = _time_strip_ampm(booking_time)
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
        "bookingFrom": "Online",
        "bookingMode": "Online",
        "processType": "Bio",
        "isTaxExempted": "N",
        "applId": resched_to.get("applId"),
        "appointmentDate": booking_date,
        "appointmentType": resched_to.get("appointmentType", "Normal"),
        "slotQuotaSequenceNumber": slot_quota_seq_no,
        "newVscId": target_vsc["vscId"],
        "numberOfApplicants": resched_to.get("numberOfApplicants", 1),
        "startTime": start_time,
        "feeTOs": None,
        "languageName": "English",
        "workingDays": None,
        "workingHours": None,
        "telephone": None,
        "captchaTO": captcha_payload,
    }


# ── Captcha auto-collection ───────────────────────────────────────────────────

_CAPTCHA_DIR = os.path.join(os.path.dirname(__file__), "captcha_solver", "real_captchas")
_RETRAIN_THRESHOLD = 50  # retrain after this many new captchas collected this session
_new_captcha_count = 0


def _save_captcha(img_bytes: bytes, label: str):
    global _new_captcha_count
    try:
        os.makedirs(_CAPTCHA_DIR, exist_ok=True)
        existing = [f for f in os.listdir(_CAPTCHA_DIR) if f.endswith(".png")]
        idx = len(existing) + 1
        label_clean = label.strip().upper()[:5]
        filename = f"{label_clean}_{idx:04d}.png"
        path = os.path.join(_CAPTCHA_DIR, filename)
        with open(path, "wb") as f:
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
        print("[RETRAIN] train.py not found — skipping")
        return
    import subprocess
    python = os.path.join(os.path.dirname(__file__), "venv312", "Scripts", "python.exe")
    print(f"[RETRAIN] Starting retraining ({_new_captcha_count} new captchas collected)...")
    subprocess.Popen([python, train_script], cwd=os.path.dirname(__file__))
    print(f"[RETRAIN] Retraining launched in background")


# ── Discord notifications ─────────────────────────────────────────────────────

def _notify_discord(message: str, webhook: str = None):
    url = webhook or DISCORD_WEBHOOK
    if not url:
        return
    try:
        requests.post(url, json={"content": message}, timeout=10)
    except Exception as e:
        print(f"[DISCORD] Failed to send notification: {e}")


# ── Alert on success ──────────────────────────────────────────────────────────

def _play_alert():
    siren = os.path.join(os.path.dirname(__file__), "siren.wav")
    if os.path.exists(siren):
        for _ in range(5):
            winsound.PlaySound(siren, winsound.SND_FILENAME)
    else:
        for _ in range(10):
            winsound.Beep(1000, 300)
            time.sleep(0.1)


# ── Date filtering ───────────────────────────────────────────────────────��────


def _is_before_urgent(date_str: str) -> bool:
    """Return True if date_str is strictly before URGENT_MEDICAL_DATE (or no urgent date set)."""
    if not URGENT_MEDICAL_DATE:
        return True
    try:
        return datetime.strptime(date_str, "%Y-%m-%d") < datetime.strptime(URGENT_MEDICAL_DATE, "%Y-%m-%d")
    except ValueError:
        return True

# ── Main reschedule loop ──────────────────────────────────────────────────────

def _get_year_months() -> list[tuple[int, int]]:
    """Convert MONTHS_TO_CHECK to (year, month) pairs relative to now."""
    now = datetime.now()
    result = []
    for name in MONTHS_TO_CHECK:
        m = MONTH_NAMES.get(name)
        if m is None:
            print(f"[WARN] Unknown month: {name!r}")
            continue
        y = now.year
        # If the month has already passed this year, check next year
        if m < now.month:
            y += 1
        result.append((y, m))
    return result


def run():
    _load_proxies()

    # Pre-compute encoded passport/visa (deterministic — same key+IV always)
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

            # ── Steps 2+3: Get captcha + validate (retry up to 12 times) ─────
            resched_to = None
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
                print("[API] Calling rescheduleDetailValidateCaptcha...")
                try:
                    resched_to = reschedule_validate_captcha(
                        sess, token, captcha_to, captcha_text, enc_visa, enc_pass)
                    _save_captcha(img_bytes, captcha_text)
                    break
                except TokenActiveError:
                    # Wrong captcha created a session — clear it before next attempt
                    print(f"[E013] Session created by wrong captcha — clearing...")
                    time.sleep(1)
                    delete_old_token(sess, token)
                    continue
                except CaptchaError as e:
                    print(f"[CAPTCHA1] {e} — getting new captcha...")
                    time.sleep(1)
                    continue
            if resched_to is None:
                raise ApiError("Failed to pass captcha+validation after 8 attempts")

            print(f"[RESCHED] applId={resched_to.get('applId')}  "
                  f"ref={resched_to.get('applicationReferenceNumber')}  "
                  f"visaTypeId={resched_to.get('visaTypeId')}")

            visa_type_id = resched_to.get("visaTypeId", 51)
            country_id   = 3050
            applicants   = resched_to.get("applicant") or []
            sponsor_type_ids = list({
                str(a.get("sponsorTypeCode", "2"))
                for a in applicants
                if a.get("sponsorTypeCode")
            }) or ["2"]

            # ── Step 4: VSC details (optional) ────────────────────────────────
            try:
                get_vsc_details(sess, token, country_id, visa_type_id)
                print("[API] getVscDetails OK")
            except Exception as e:
                print(f"[API] getVscDetails skipped: {e}")

            # ── Poll loop — reuse same session until token expires ────────────
            poll_num = 0
            token_start = time.time()
            TOKEN_LIFETIME = 18 * 60  # refresh after 18 min (token lasts ~20 min)

            while time.time() - token_start < TOKEN_LIFETIME:
                poll_num += 1
                print(f"\n[POLL {poll_num}] {datetime.now().strftime('%H:%M:%S')}")

                # ── Phase 1: scan all months, collect every date+slot ─────────
                urgent_slots = []
                normal_slots = []
                got_429     = False

                for (year, month) in year_months:
                    print(f"[SCAN] {calendar.month_name[month]} {year} → {QVC_CENTER}")
                    try:
                        available_dates = get_appointment_dates(
                            sess, token, year, month,
                            target_vsc["vscId"], country_id, visa_type_id,
                            enc_visa, enc_pass, sponsor_type_ids,
                        )
                    except SessionExpiredError as e:
                        print(f"[SESSION] Server session expired ({e}) — restarting full session...")
                        raise SessionExpiredError(e)
                    except RateLimitError:
                        print(f"[429] Rate limited — switching proxy, keeping token...")
                        proxy_line = _pick_proxy()
                        sess = _make_session(proxy_line)
                        got_429 = True
                        break

                    if not available_dates:
                        print(f"[SCAN] No dates in {calendar.month_name[month]}")
                        continue

                    for date_str in available_dates:
                        _is_urgent = _is_before_urgent(date_str)
                        _label = "URGENT" if (URGENT_MEDICAL_DATE and _is_urgent) else ("NORMAL" if URGENT_MEDICAL_DATE else "")
                        _tag   = f"[{_label}] " if _label else ""
                        slots = fetch_slots(
                            sess, token, date_str,
                            target_vsc["vscId"], country_id,
                            VISA_NUMBER, PASSPORT_NUMBER,
                            enc_visa, enc_pass, sponsor_type_ids,
                        )
                        _dt = datetime.strptime(date_str, "%Y-%m-%d")
                        _month_name   = _dt.strftime("%B")
                        _date_display = f"{_dt.month}-{_dt.day}-{_dt.year}"
                        if not slots:
                            print(f"{_tag}Month      : {_month_name}")
                            print(f"{_tag}Date       : {_date_display}")
                            print(f"{_tag}Time Slots : Not Available")
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
                                    urgent_slots.append((date_str, stime, etime, slot_id, avail))
                                else:
                                    normal_slots.append((date_str, stime, etime, slot_id, avail))
                                _avail_times.append(stime.replace(" ", ""))
                        print(f"{_tag}Month      : {_month_name}")
                        print(f"{_tag}Date       : {_date_display}")
                        if _avail_times:
                            print(f"{_tag}Time Slots : {'  '.join(_avail_times)}")
                        else:
                            print(f"{_tag}Time Slots : Not Available")

                # ── Phase 2: print full summary ───────────────────────────────
                all_slots = urgent_slots + normal_slots
                if not all_slots:
                    print(f"[POLL] No slots found. Next scan in {POLL_INTERVAL}s...")
                    time.sleep(POLL_INTERVAL)
                    continue

                urgent_slots.sort(key=lambda x: (x[0], x[1]))
                normal_slots.sort(key=lambda x: (x[0], x[1]))

                print(f"\n{'─'*50}")
                if urgent_slots:
                    print(f"[URGENT] {len(urgent_slots)} slot(s) at {QVC_CENTER}:")
                    for (ds, st, et, sid, av) in urgent_slots:
                        _dtu = datetime.strptime(ds, "%Y-%m-%d")
                        print(f"[URGENT] Month: {_dtu.strftime('%B')}  Date: {_dtu.month}-{_dtu.day}-{_dtu.year}  Time: {st.replace(' ','')}")
                if normal_slots:
                    print(f"[NORMAL] {len(normal_slots)} slot(s) at {QVC_CENTER}:")
                    for (ds, st, et, sid, av) in normal_slots:
                        _dtn = datetime.strptime(ds, "%Y-%m-%d")
                        print(f"[NORMAL] Month: {_dtn.strftime('%B')}  Date: {_dtn.month}-{_dtn.day}-{_dtn.year}  Time: {st.replace(' ','')}")
                print(f"{'─'*50}")

                # Discord alert
                _dmsg = "\U0001f514 **SLOTS AVAILABLE**\n"
                _dmsg += f"\U0001f4cd **{QVC_CENTER}**\n\n"
                if urgent_slots:
                    _dmsg += "\U0001f534 **URGENT** *(before " + (URGENT_MEDICAL_DATE or "") + ")*\n"
                    _u_dates = {}
                    for (ds, st, et, sid, av) in urgent_slots:
                        _u_dates.setdefault(ds, []).append(st.replace(" ", ""))
                    for ds, times in _u_dates.items():
                        _dtu2 = datetime.strptime(ds, "%Y-%m-%d")
                        _dmsg += f"\U0001f4c5  **{_dtu2.strftime('%B')}**  |  {_dtu2.month}-{_dtu2.day}-{_dtu2.year}\n"
                        _dmsg += "\u23f0  " + "   \u2022   ".join(times) + "\n\n"
                if normal_slots:
                    _dmsg += "\U0001f7e2 **NORMAL**\n"
                    _n_dates = {}
                    for (ds, st, et, sid, av) in normal_slots:
                        _n_dates.setdefault(ds, []).append(st.replace(" ", ""))
                    for ds, times in _n_dates.items():
                        _dtn2 = datetime.strptime(ds, "%Y-%m-%d")
                        _dmsg += f"\U0001f4c5  **{_dtn2.strftime('%B')}**  |  {_dtn2.month}-{_dtn2.day}-{_dtn2.year}\n"
                        _dmsg += "\u23f0  " + "   \u2022   ".join(times) + "\n\n"
                # Discord alerts — urgent and normal completely separate
                if urgent_slots:
                    _umsg = "\U0001f514 **URGENT RESCHEDULE SLOTS AVAILABLE**\n"
                    _umsg += f"\U0001f4cd **{QVC_CENTER}**\n"
                    _umsg += "\U0001f534 Before: " + URGENT_MEDICAL_DATE + "\n\n"
                    _u_d = {}
                    for (ds, st, et, sid, av) in urgent_slots:
                        _u_d.setdefault(ds, []).append(st.replace(" ", ""))
                    for ds, times in _u_d.items():
                        _udt = datetime.strptime(ds, "%Y-%m-%d")
                        _umsg += f"\U0001f4c5  **{_udt.strftime('%B')}**  |  {_udt.month}-{_udt.day}-{_udt.year}\n"
                        _umsg += "⏰  " + "   •   ".join(times) + "\n\n"
                    _notify_discord(_umsg.strip(), webhook=DISCORD_URGENT_WEBHOOK)
                if normal_slots:
                    _nmsg = "\U0001f514 **RESCHEDULE SLOTS AVAILABLE**\n"
                    _nmsg += f"\U0001f4cd **{QVC_CENTER}**\n\n"
                    _n_d = {}
                    for (ds, st, et, sid, av) in normal_slots:
                        _n_d.setdefault(ds, []).append(st.replace(" ", ""))
                    for ds, times in _n_d.items():
                        _ndt = datetime.strptime(ds, "%Y-%m-%d")
                        _nmsg += f"\U0001f4c5  **{_ndt.strftime('%B')}**  |  {_ndt.month}-{_ndt.day}-{_ndt.year}\n"
                        _nmsg += "⏰  " + "   •   ".join(times) + "\n\n"
                    _notify_discord(_nmsg.strip(), webhook=DISCORD_WEBHOOK)

                # Decide what to book
                slots_to_book = urgent_slots if URGENT_MEDICAL_DATE else all_slots
                if not slots_to_book:
                    print(f"[POLL] No URGENT slots found before {URGENT_MEDICAL_DATE}. Waiting {POLL_INTERVAL}s...")
                    time.sleep(POLL_INTERVAL)
                    continue
                _eb2 = slots_to_book[0]
                print(f"[BOOKING] Attempting earliest: {_eb2[0]}  {_eb2[1]}")

                found = False
                for (date_str, start_time, end_time, slot_id, _avail) in slots_to_book:
                    booking_time = _time_strip_ampm(start_time)
                    schedule_to  = _build_schedule_to(resched_to, target_vsc, date_str, booking_time)

                    try:
                        uid = check_slot_available(sess, token, schedule_to, slot_id, enc_visa, enc_pass)
                    except ApiError as e:
                        print(f"[SLOT] checkslotavailable error: {e}")
                        continue
                    if uid is None:
                        print(f"[SLOT] {date_str} {start_time} — no longer available, trying next...")
                        continue
                    print(f"[SLOT] Confirmed: {date_str} {start_time}  uid={uid}")

                    # Captcha 2
                    print("[CAPTCHA2] Getting schedule captcha...")
                    sch_cap_to = None
                    cap2_valid = False
                    for cap2_attempt in range(15):
                        sch_cap_to = get_schedule_captcha(
                            sess, token, enc_visa, enc_pass,
                            sch_cap_to.get("captchaId") if sch_cap_to else None,
                        )
                        img2_b64 = sch_cap_to.get("imageString", "")
                        if not img2_b64:
                            break
                        cap2_text = solve_captcha(base64.b64decode(img2_b64))
                        if not cap2_text:
                            print(f"[CAPTCHA2] Could not solve ({cap2_attempt+1}/15), refreshing...")
                            continue
                        print(f"[CAPTCHA2] Answer: '{cap2_text}' ({cap2_attempt+1}/15)")
                        try:
                            cap2_valid = validate_schedule_captcha(
                                sess, token, enc_visa, enc_pass, sch_cap_to["captchaId"], cap2_text)
                        except ApiError as e:
                            print(f"[CAPTCHA2] Validation error: {e}")
                        if cap2_valid:
                            break
                        print(f"[CAPTCHA2] Wrong answer...")
                        if (cap2_attempt + 1) % 5 == 0:
                            print(f"[CAPTCHA2] 5 wrong in a row — waiting 10s...")
                            time.sleep(10)

                    if not cap2_valid:
                        print(f"[CAPTCHA2] Failed after 15 attempts — still attempting save")

                    save_payload = _build_save_payload(resched_to, target_vsc, date_str, booking_time, uid,
                                                       captcha_to=sch_cap_to, captcha_value=cap2_text)
                    if DRY_RUN:
                        print(f"[DRY_RUN] Would save: {date_str} {booking_time} at {QVC_CENTER}  uid={uid}")
                        print(f"[DRY_RUN] Payload: {json.dumps(save_payload, indent=2)}")
                        found = True
                        break

                    print(f"[SAVE] Saving reschedule for {date_str} {booking_time}...")
                    try:
                        result = save_reschedule(sess, token, save_payload)
                    except ApiError as e:
                        print(f"[SAVE] Error: {e}")
                        continue

                    msg_code = result.get("messageCode", "")
                    status   = result.get("statusCode", "")
                    icr      = result.get("icrResponse") or {}

                    if msg_code == "E016":
                        print("[SAVE] E016 captcha wrong — retrying next slot")
                        continue
                    if msg_code:
                        print(f"[SAVE] Server error: {msg_code}  status={status}")
                        continue

                    if status in ("OK", "200 OK") or icr.get("paymentStatus") == "SUCCESS":
                        print("\n" + "★" * 60)
                        print(f"✓  RESCHEDULED!  {date_str} at {booking_time}")
                        print(f"   Center: {QVC_CENTER}")
                        print(f"   ref:    {resched_to.get('applicationReferenceNumber')}")
                        print("★" * 60 + "\n")
                        _play_alert()
                        found = True
                        break

                if found:
                    print("[DONE] Appointment rescheduled successfully. Exiting.")
                    return

                print(f"[POLL] No slots. Next scan in {POLL_INTERVAL}s...")
                time.sleep(POLL_INTERVAL)

            print(f"[SESSION] Token nearing expiry — refreshing session...")

        except RateLimitError as e:
            print(f"[429] {e} — switching proxy and retrying in 5s...")
            proxy_line = _pick_proxy()
            sess = _make_session(proxy_line)
            time.sleep(5)
        except SessionExpiredError:
            print(f"[SESSION] Restarting session immediately...")
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
