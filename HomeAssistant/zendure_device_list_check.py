import ast
import base64
import binascii
import hashlib
import json
import secrets
import sys
import time
import traceback
import urllib.request
from datetime import datetime

CONST_FILE = "/config/custom_components/zendure_ha/const.py"
CONFIG_ENTRIES = "/config/.storage/core.config_entries"
OUTPUT_FILE = "/config/custom_components/zendure_ha/zendure_device_list_output.txt"

def write_output(lines):
    text = "\n".join(lines) + "\n"
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        f.write(text)
    print("Output written to:", OUTPUT_FILE)

def read_string_constants_from_file(path):
    """
    Parse const.py without importing it.
    This avoids the SSH add-on problem: No module named 'homeassistant'.
    """
    with open(path, "r", encoding="utf-8") as f:
        source = f.read()

    tree = ast.parse(source, filename=path)
    constants = {}

    for node in tree.body:
        if isinstance(node, ast.Assign):
            value = None

            if isinstance(node.value, ast.Constant) and isinstance(node.value.value, str):
                value = node.value.value

            if value is None:
                continue

            for target in node.targets:
                if isinstance(target, ast.Name):
                    constants[target.id] = value

    return constants

lines = []
lines.append("Zendure deviceList API check")
lines.append("Generated: " + datetime.now().isoformat(timespec="seconds"))
lines.append("=" * 60)

try:
    consts = read_string_constants_from_file(CONST_FILE)

    conf_apptoken = consts.get("CONF_APPTOKEN")
    conf_hakey = consts.get("CONF_HAKEY")

    lines.append("CONST_FILE: " + CONST_FILE)
    lines.append("CONF_APPTOKEN key: " + str(conf_apptoken))
    lines.append("CONF_HAKEY found: " + str(bool(conf_hakey)))

    if not conf_apptoken:
        lines.append("")
        lines.append("Available CONF_* constants:")
        for key in sorted(k for k in consts if k.startswith("CONF_")):
            lines.append("  - " + key)
        raise RuntimeError("CONF_APPTOKEN not found in const.py")

    if not conf_hakey:
        lines.append("")
        lines.append("Available CONF_* constants:")
        for key in sorted(k for k in consts if k.startswith("CONF_")):
            lines.append("  - " + key)
        raise RuntimeError("CONF_HAKEY not found in const.py")

except Exception as e:
    lines.append("ERROR: Could not read Zendure constants from const.py")
    lines.append(repr(e))
    lines.append(traceback.format_exc())
    write_output(lines)
    sys.exit(1)

try:
    with open(CONFIG_ENTRIES, "r", encoding="utf-8") as f:
        cfg = json.load(f)
except Exception as e:
    lines.append("ERROR: Could not read Home Assistant config entries")
    lines.append(CONFIG_ENTRIES)
    lines.append(repr(e))
    lines.append(traceback.format_exc())
    write_output(lines)
    sys.exit(1)

entries = cfg.get("data", {}).get("entries", [])

zendure_entries = [
    e for e in entries
    if e.get("domain") in ("zendure_ha", "zendure")
]

lines.append("Zendure config entries found: " + str(len(zendure_entries)))

if not zendure_entries:
    lines.append("ERROR: No Zendure config entry found.")
    write_output(lines)
    sys.exit(1)

entry = zendure_entries[0]
entry_data = entry.get("data", {})

lines.append("Using config entry domain: " + str(entry.get("domain")))
lines.append("Using config entry title: " + str(entry.get("title")))

token = entry_data.get(conf_apptoken)

if not token:
    lines.append("ERROR: No app token found with key: " + str(conf_apptoken))
    lines.append("Available config entry data keys:")
    for key in sorted(entry_data.keys()):
        lines.append("  - " + str(key))
    write_output(lines)
    sys.exit(1)

try:
    decoded = base64.b64decode(str(token)).decode("utf-8")
    api_url, app_key = decoded.rsplit(".", 1)
except (binascii.Error, UnicodeDecodeError, ValueError) as e:
    lines.append("ERROR: Token decode failed")
    lines.append(repr(e))
    lines.append(traceback.format_exc())
    write_output(lines)
    sys.exit(1)

timestamp = int(time.time())
nonce = str(secrets.randbelow(90000) + 10000)

body = {
    "appKey": app_key
}

sign_params = {
    **body,
    "timestamp": timestamp,
    "nonce": nonce,
}

body_str = "".join(f"{k}{v}" for k, v in sorted(sign_params.items()))
sign_str = f"{conf_hakey}{body_str}{conf_hakey}"
sign = hashlib.sha1(sign_str.encode("utf-8")).hexdigest().upper()

headers = {
    "Content-Type": "application/json",
    "timestamp": str(timestamp),
    "nonce": nonce,
    "clientid": "zenHa",
    "sign": sign,
}

url = f"{api_url}/api/ha/deviceList"

lines.append("")
lines.append("Calling Zendure API:")
lines.append(url)
lines.append("")

req = urllib.request.Request(
    url,
    data=json.dumps(body).encode("utf-8"),
    headers=headers,
    method="POST",
)

try:
    with urllib.request.urlopen(req, timeout=20) as resp:
        raw = resp.read().decode("utf-8")
        status = resp.status
except Exception as e:
    lines.append("ERROR: Request failed")
    lines.append(repr(e))

    # If urllib raises HTTPError, the API may still return a useful response body.
    if hasattr(e, "read"):
        try:
            err_body = e.read().decode("utf-8", errors="replace")
            lines.append("")
            lines.append("HTTP error response body:")
            lines.append(err_body)
        except Exception as read_err:
            lines.append("Could not read HTTP error body:")
            lines.append(repr(read_err))

    lines.append("")
    lines.append(traceback.format_exc())
    write_output(lines)
    sys.exit(1)

lines.append("HTTP status: " + str(status))
lines.append("Raw response length: " + str(len(raw)))

try:
    result = json.loads(raw)
except Exception as e:
    lines.append("ERROR: JSON parse failed")
    lines.append(repr(e))
    lines.append("")
    lines.append("Raw response first 2000 chars:")
    lines.append(raw[:2000])
    write_output(lines)
    sys.exit(1)

lines.append("API code: " + str(result.get("code")))
lines.append("API success: " + str(result.get("success")))
lines.append("API msg: " + str(result.get("msg")))
lines.append("")

api_data = result.get("data", {})
device_list = api_data.get("deviceList", [])

lines.append("Devices from Zendure /api/ha/deviceList")
lines.append("-" * 60)

for idx, dev in enumerate(device_list, start=1):
    if not isinstance(dev, dict):
        lines.append(f"{idx}. Invalid device entry type: {type(dev).__name__}")
        continue

    name = (
        dev.get("name")
        or dev.get("deviceName")
        or dev.get("alias")
        or dev.get("productName")
        or dev.get("productModel")
        or ""
    )

    sn = str(dev.get("sn", "") or dev.get("snNumber", ""))
    ip = dev.get("ip")
    product = dev.get("productModel") or dev.get("product") or ""
    device_id = str(dev.get("deviceId", ""))

    lines.append(f"{idx}. name      : {name}")
    lines.append(f"   product   : {product}")
    lines.append(f"   sn_end    : {sn[-4:] if sn else None}")
    lines.append(f"   device_id : {device_id[-6:] if device_id else None}")
    lines.append(f"   ip        : {ip}")
    lines.append("")

lines.append("-" * 60)
lines.append("Check complete.")
lines.append("")
lines.append("Note: Full serial numbers and tokens were not written to this file.")

write_output(lines)
