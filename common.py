"""Gemeinsame Funktionen: Config laden, Token-Management."""

import glob
import json
import os
import time
import threading
import urllib.parse
import urllib.request

CONFIG_DIR = os.path.expanduser("~/.config/outlook-proxy")
CONFIG_FILE = os.path.join(CONFIG_DIR, "config.json")
TOKEN_FILE = os.path.join(CONFIG_DIR, "token.json")

_token_lock = threading.Lock()
_config_cache = None


def load_config():
    global _config_cache
    if _config_cache:
        return _config_cache

    if not os.path.exists(CONFIG_FILE):
        print(f"Config nicht gefunden: {CONFIG_FILE}")
        print(f"Bitte config.example.json kopieren und anpassen:")
        print(f"  mkdir -p {CONFIG_DIR}")
        print(f"  cp config.example.json {CONFIG_FILE}")
        raise SystemExit(1)

    with open(CONFIG_FILE) as f:
        _config_cache = json.load(f)
    return _config_cache


def token_url():
    cfg = load_config()
    return f"https://login.microsoftonline.com/{cfg['tenant_id']}/oauth2/v2.0/token"


def load_tokens():
    if os.path.exists(TOKEN_FILE):
        with open(TOKEN_FILE) as f:
            return json.load(f)
    return None


def save_tokens(tokens):
    os.makedirs(CONFIG_DIR, exist_ok=True)
    with open(TOKEN_FILE, "w") as f:
        json.dump(tokens, f, indent=2)
    os.chmod(TOKEN_FILE, 0o600)


def get_access_token(scopes=None):
    """Holt einen gültigen Access-Token, erneuert bei Bedarf."""
    cfg = load_config()

    if not scopes:
        scopes = (
            "https://outlook.office.com/IMAP.AccessAsUser.All "
            "https://outlook.office.com/SMTP.Send "
            "https://outlook.office.com/EWS.AccessAsUser.All "
            "offline_access"
        )

    with _token_lock:
        tokens = load_tokens()
        if not tokens:
            return None

        expires_at = tokens.get("expires_at", 0)
        if time.time() < expires_at - 120:
            return tokens["access_token"]

        try:
            post_data = urllib.parse.urlencode({
                "client_id": cfg["client_id"],
                "grant_type": "refresh_token",
                "refresh_token": tokens["refresh_token"],
                "scope": scopes,
                "redirect_uri": cfg["redirect_uri"],
            }).encode()

            req = urllib.request.Request(token_url(), data=post_data)
            req.add_header("Content-Type", "application/x-www-form-urlencoded")
            resp = urllib.request.urlopen(req)
            result = json.loads(resp.read())

            tokens["access_token"] = result["access_token"]
            tokens["refresh_token"] = result["refresh_token"]
            tokens["expires_at"] = time.time() + int(result.get("expires_in", 3600))
            save_tokens(tokens)
            return tokens["access_token"]
        except Exception as e:
            print(f"[token] Refresh fehlgeschlagen: {e}")
            return None


def find_thunderbird_profile():
    """Findet das aktive Thunderbird-Profil."""
    cfg = load_config()
    pattern = os.path.expanduser(cfg.get("thunderbird_profile", "~/.thunderbird/*.default-esr"))
    matches = glob.glob(pattern)
    if matches:
        return matches[0]

    # Fallback: profiles.ini parsen
    tb_dir = os.path.expanduser("~/.thunderbird")
    ini_path = os.path.join(tb_dir, "profiles.ini")
    if os.path.exists(ini_path):
        with open(ini_path) as f:
            for line in f:
                if line.startswith("Path="):
                    path = line.strip().split("=", 1)[1]
                    full = os.path.join(tb_dir, path)
                    if os.path.isdir(full):
                        return full
    return None
