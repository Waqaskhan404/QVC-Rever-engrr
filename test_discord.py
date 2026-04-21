import requests

url = "https://discordapp.com/api/webhooks/1495888024561516777/Hdcv7CY-fE8zjtxo3eYupCVLYzapg_cIlY3lbF0YSLWWyor1TIq7hBWYMCn4RsF2TOGO"
r = requests.post(url, json={"content": "test from python"}, timeout=10)
print(f"status: {r.status_code}")
print(r.text)
