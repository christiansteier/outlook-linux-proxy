#!/bin/bash
set -e

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
CONFIG_DIR="$HOME/.config/outlook-proxy"
BIN_DIR="$HOME/.local/bin"
SERVICE_DIR="$HOME/.config/systemd/user"

echo "=== Outlook Linux Proxy — Installation ==="
echo ""

# 1. Config
mkdir -p "$CONFIG_DIR"
if [ ! -f "$CONFIG_DIR/config.json" ]; then
    echo "Erstelle Config-Vorlage..."
    cp "$SCRIPT_DIR/config.example.json" "$CONFIG_DIR/config.json"
    echo "  WICHTIG: Bitte $CONFIG_DIR/config.json anpassen!"
    echo "  Mindestens tenant_id und email eintragen."
    echo ""
    read -p "Jetzt bearbeiten? (j/n) " -n 1 -r
    echo ""
    if [[ $REPLY =~ ^[Jj]$ ]]; then
        ${EDITOR:-nano} "$CONFIG_DIR/config.json"
    fi
else
    echo "Config existiert bereits: $CONFIG_DIR/config.json"
fi

# 2. Skripte installieren
echo "Installiere Skripte..."
mkdir -p "$BIN_DIR"
for f in common.py imap-smtp-proxy.py caldav-bridge.py init-token.py; do
    cp "$SCRIPT_DIR/$f" "$BIN_DIR/outlook-proxy-$f"
    # Fix imports for installed location
    sed -i "s|os.path.dirname(os.path.abspath(__file__))|\"$BIN_DIR\"|g" "$BIN_DIR/outlook-proxy-$f"
done
chmod +x "$BIN_DIR"/outlook-proxy-*.py
# common.py muss ohne Prefix erreichbar sein
cp "$SCRIPT_DIR/common.py" "$BIN_DIR/common.py"

# 3. systemd Services
echo "Installiere systemd Services..."
mkdir -p "$SERVICE_DIR"

cat > "$SERVICE_DIR/outlook-mail-proxy.service" << EOF
[Unit]
Description=Outlook IMAP/SMTP Proxy (OAuth2)
After=network-online.target
Wants=network-online.target

[Service]
Type=simple
ExecStart=/usr/bin/python3 $BIN_DIR/outlook-proxy-imap-smtp-proxy.py
Restart=on-failure
RestartSec=10

[Install]
WantedBy=default.target
EOF

cat > "$SERVICE_DIR/outlook-caldav-bridge.service" << EOF
[Unit]
Description=Outlook CalDAV Bridge (EWS/OAuth2)
After=network-online.target
Wants=network-online.target

[Service]
Type=simple
ExecStart=/usr/bin/python3 $BIN_DIR/outlook-proxy-caldav-bridge.py
Restart=on-failure
RestartSec=10

[Install]
WantedBy=default.target
EOF

systemctl --user daemon-reload

# 4. Token initialisieren
if [ ! -f "$CONFIG_DIR/token.json" ]; then
    echo ""
    echo "Token initialisieren (Device Code Flow im Browser)..."
    python3 "$BIN_DIR/outlook-proxy-init-token.py"
fi

# 5. Services starten
echo ""
echo "Starte Services..."
systemctl --user enable --now outlook-mail-proxy.service
systemctl --user enable --now outlook-caldav-bridge.service
loginctl enable-linger "$USER" 2>/dev/null || true

sleep 3
echo ""
echo "=== Status ==="
systemctl --user status outlook-mail-proxy.service --no-pager | head -5
systemctl --user status outlook-caldav-bridge.service --no-pager | head -5

echo ""
echo "=== Fertig! ==="
echo ""
echo "Evolution-Einstellungen:"
echo "  IMAP:    localhost:1143 (keine Verschlüsselung, Passwort-Auth)"
echo "  SMTP:    localhost:1025 (keine Verschlüsselung, Passwort-Auth)"
echo "  CalDAV:  http://localhost:1080/calendar/"
echo ""
echo "Befehle:"
echo "  Status:    systemctl --user status outlook-mail-proxy outlook-caldav-bridge"
echo "  Neustart:  systemctl --user restart outlook-mail-proxy outlook-caldav-bridge"
echo "  Logs:      journalctl --user -u outlook-mail-proxy -f"
echo "  Token neu: python3 $BIN_DIR/outlook-proxy-init-token.py"
