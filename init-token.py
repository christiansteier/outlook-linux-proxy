#!/usr/bin/env python3
"""
Initialisiert den OAuth2-Token für die Outlook-Proxys.

Standard: Device Code Flow — öffne die angezeigte URL im Browser,
gib den Code ein und melde dich an. Kein Thunderbird nötig.

Fallback: --thunderbird extrahiert den Token aus einer bestehenden
Thunderbird-Installation (für den Fall, dass der Device Code Flow
im Tenant blockiert wird).
"""

import argparse
import base64
import ctypes
import ctypes.util
import json
import os
import sys
import time
import urllib.parse
import urllib.request

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import load_config, save_tokens, find_thunderbird_profile, token_url


def device_code_flow():
    """OAuth2 Device Code Flow — funktioniert ohne Thunderbird."""
    cfg = load_config()

    scopes = (
        "https://outlook.office.com/IMAP.AccessAsUser.All "
        "https://outlook.office.com/SMTP.Send "
        "https://outlook.office.com/EWS.AccessAsUser.All "
        "offline_access"
    )

    # 1. Device Code anfordern
    post_data = urllib.parse.urlencode({
        "client_id": cfg["client_id"],
        "scope": scopes,
    }).encode()

    req = urllib.request.Request(
        f"https://login.microsoftonline.com/{cfg['tenant_id']}/oauth2/v2.0/devicecode",
        data=post_data,
    )
    req.add_header("Content-Type", "application/x-www-form-urlencoded")

    try:
        resp = urllib.request.urlopen(req)
    except urllib.error.HTTPError as e:
        body = json.loads(e.read().decode())
        print(f"FEHLER: {body.get('error_description', str(e))}")
        sys.exit(1)

    result = json.loads(resp.read())
    device_code = result["device_code"]
    interval = result.get("interval", 5)

    print()
    print("=" * 60)
    print(f"  Öffne im Browser:  {result['verification_uri']}")
    print(f"  Gib diesen Code ein:  {result['user_code']}")
    print("=" * 60)
    print()
    print("Warte auf Anmeldung...")

    # 2. Pollen bis der User sich angemeldet hat
    poll_data = urllib.parse.urlencode({
        "client_id": cfg["client_id"],
        "grant_type": "urn:ietf:params:oauth:grant-type:device_code",
        "device_code": device_code,
    }).encode()

    while True:
        time.sleep(interval)
        req = urllib.request.Request(token_url(), data=poll_data)
        req.add_header("Content-Type", "application/x-www-form-urlencoded")

        try:
            resp = urllib.request.urlopen(req)
            tokens = json.loads(resp.read())
            return tokens
        except urllib.error.HTTPError as e:
            body = json.loads(e.read().decode())
            error = body.get("error", "")
            if error == "authorization_pending":
                continue
            elif error == "slow_down":
                interval += 5
                continue
            elif error == "expired_token":
                print("FEHLER: Code abgelaufen. Bitte nochmal starten.")
                sys.exit(1)
            else:
                print(f"FEHLER: {body.get('error_description', error)}")
                sys.exit(1)


def thunderbird_flow():
    """Fallback: Token aus Thunderbird extrahieren."""
    profile = find_thunderbird_profile()
    if not profile:
        print("FEHLER: Kein Thunderbird-Profil gefunden.")
        sys.exit(1)
    print(f"  Profil: {profile}")

    libnss_path = None
    for p in ["/usr/lib/x86_64-linux-gnu/libnss3.so", ctypes.util.find_library("nss3")]:
        if p and os.path.exists(p):
            libnss_path = p
            break
    if not libnss_path:
        print("FEHLER: libnss3 nicht gefunden.")
        sys.exit(1)

    libnss = ctypes.CDLL(libnss_path)

    class SECItem(ctypes.Structure):
        _fields_ = [("type", ctypes.c_uint), ("data", ctypes.c_char_p), ("len", ctypes.c_uint)]

    ret = libnss.NSS_Init(profile.encode())
    if ret != 0:
        print(f"FEHLER: NSS_Init fehlgeschlagen")
        sys.exit(1)

    PK11_PW_CB = ctypes.CFUNCTYPE(ctypes.c_char_p, ctypes.c_void_p, ctypes.c_int, ctypes.c_void_p)
    libnss.PK11_SetPasswordFunc(PK11_PW_CB(lambda s, r, a: b""))

    def decrypt(enc_b64):
        enc = base64.b64decode(enc_b64)
        inp = SECItem(0, enc, len(enc))
        out = SECItem(0, None, 0)
        if libnss.PK11SDR_Decrypt(ctypes.byref(inp), ctypes.byref(out), None) == 0 and out.data:
            return ctypes.string_at(out.data, out.len).decode("utf-8")
        return None

    logins_path = os.path.join(profile, "logins.json")
    with open(logins_path) as f:
        data = json.load(f)

    username = refresh_token = None
    for login in data.get("logins", []):
        if login.get("hostname") == "oauth://login.microsoftonline.com":
            username = decrypt(login["encryptedUsername"])
            refresh_token = decrypt(login["encryptedPassword"])
            break

    libnss.NSS_Shutdown()

    if not refresh_token:
        print("FEHLER: Kein OAuth2-Token in Thunderbird gefunden.")
        sys.exit(1)

    print(f"  Benutzer: {username}")

    # Refresh-Token tauschen
    cfg = load_config()
    scopes = (
        "https://outlook.office.com/IMAP.AccessAsUser.All "
        "https://outlook.office.com/SMTP.Send "
        "https://outlook.office.com/EWS.AccessAsUser.All "
        "offline_access"
    )
    post_data = urllib.parse.urlencode({
        "client_id": cfg["client_id"],
        "grant_type": "refresh_token",
        "refresh_token": refresh_token,
        "scope": scopes,
        "redirect_uri": cfg["redirect_uri"],
    }).encode()

    req = urllib.request.Request(token_url(), data=post_data)
    req.add_header("Content-Type", "application/x-www-form-urlencoded")

    try:
        resp = urllib.request.urlopen(req)
        return json.loads(resp.read())
    except urllib.error.HTTPError as e:
        print(f"FEHLER: Token-Austausch fehlgeschlagen: {e.read().decode()[:200]}")
        sys.exit(1)


def main():
    parser = argparse.ArgumentParser(
        description="OAuth2-Token für Outlook-Proxy initialisieren"
    )
    parser.add_argument(
        "--thunderbird",
        action="store_true",
        help="Token aus Thunderbird extrahieren statt Device Code Flow",
    )
    args = parser.parse_args()

    cfg = load_config()

    if args.thunderbird:
        print("Token aus Thunderbird extrahieren...")
        tokens = thunderbird_flow()
    else:
        print("Device Code Flow starten...")
        tokens = device_code_flow()

    token_data = {
        "access_token": tokens["access_token"],
        "refresh_token": tokens["refresh_token"],
        "expires_at": time.time() + int(tokens.get("expires_in", 3600)),
        "user": cfg["email"],
    }

    save_tokens(token_data)
    print()
    print(f"Token gespeichert (gültig für {tokens.get('expires_in', '?')}s)")
    print()
    print("Proxys neu starten:")
    print("  systemctl --user restart outlook-mail-proxy outlook-caldav-bridge")


if __name__ == "__main__":
    main()
