# Outlook unter PopOS mit GNOME Evolution


## Das Problem

An unserer Hochschule läuft Mail und Kalender über Microsoft 365. Unter Windows und macOS mit Outlook kein Problem. Jedoch möchte ich gern unter GNUlinux GNOME Evolution als Mailclient nutzen — Mail, Kalender, Meeting-Einladungen annehmen, das volle Programm. Was ich stattdessen bekam:

```
┌─────────────────────────────────────────────┐
│  GNOME Evolution                            │
│  → OAuth2 Login bei Microsoft               │
│  → "Administratorgenehmigung erforderlich"  │
│  → GNOME Evolution: nicht überprüft         │
│  → Zugriff verweigert.                      │
└─────────────────────────────────────────────┘
```

Unsere Campus-IT hat im Azure-AD-Tenant nur bestimmte Apps freigegeben. Thunderbird ja, aber GNOME Evolution leider nicht. Auch auf Anfrage.

Mein Rechner läuft mit Pop!_OS und dem COSMIC Desktop. Da gibt es kein GNOME Online Accounts (`gnome-control-center` startet nur unter GNOME/Unity), also auch keinen zentralen OAuth2-Login für alle Apps.

**Ziel:** Mail + Kalender + Meeting-Einladungen in GNOME Evolution

---

## Was ich alles probiert habe

### Versuch 1: GNOME Online Accounts

```bash
$ gnome-control-center online-accounts
Running gnome-control-center is only supported under GNOME and Unity, exiting
```

Geht nicht unter COSMIC. Nächster Versuch.

### Versuch 2: Evolution-EWS direkt

Evolution hat ein Plugin für Exchange Web Services (`evolution-ews`). Installiert, Microsoft-365-Konto eingerichtet, Typ "Exchange Web Services", OAuth2 gewählt. Browser öffnet sich, HSD-Login erscheint, dann:

```
AADSTS700016: Application with identifier '...' was not found
in the directory 'Hochschule Düsseldorf'.
```

Evolutions Client-ID ist im HSD-Tenant nicht registriert. Sackgasse.

### Versuch 3: DavMail

[DavMail](https://davmail.sourceforge.net/) ist der Klassiker für Exchange unter Linux — ein lokaler Gateway, der IMAP/SMTP/CalDAV auf Exchange übersetzt. Installiert (`sudo apt install davmail`), konfiguriert und dann wurde es spannend.

DavMail hat vier Modi für Office 365:

| Modus | Was es macht | Was passiert |
|---|---|---|
| `O365Interactive` | Eingebetteter Browser (SWT) | `O365Interactive is not compatible with SWT` — COSMIC hat kein SWT |
| `O365Manual` | Swing-Dialog für manuelle Code-Eingabe | Dialog öffnet sich nicht unter COSMIC |
| `O365` | NTLM/Basic-Auth Fallback | `Authentication failed: invalid user or password` — Microsoft hat Basic-Auth abgeschaltet |
| `O365StoredTokenAuthenticator` | Gespeicherter Refresh-Token (undokumentiert!) | Funktioniert... einmal. |

Der letzte Modus war vielversprechend. Ich hab Thunderbirds OAuth2-Token extrahiert, in DavMails-Config geschrieben und die erste Verbindung klappte auch:

```
a1 OK Authenticated    ← Ja!
```

Dann versuchte ich eine zweite Verbindung:

```
AADSTS9002313: Invalid request. Request is malformed or invalid.
```

Warum das? Microsoft rotiert den Refresh-Token bei jeder Nutzung. DavMail bekommt einen neuen Token zurück, speichert ihn verschlüsselt und versucht beim nächsten Mal, ihn zu erneuern. Aber DavMails `O365Token`-Klasse baut die Refresh-Anfrage für mich nicht korrekt zusammen: es nutzt den v1-OAuth2-Endpoint mit unvollständigen Parametern. Ein manueller Refresh mit Python gegen denselben Endpoint (v2.0, gleiche Parameter) funktioniert dagegen einwandfrei.

**Erkenntnis für mich: DavMail funktioniert in dieser Konstellation für mich nicht!** -> Deinstalliert.

### Versuch 4: Thunderbirds Client-ID direkt

Nächste Idee: Wenn Thunderbird sich anmelden kann, nehmen wir einfach Thunderbirds Client-ID (`9e5f94bc-e8a4-4e73-b8be-63364c29d753`) und machen unseren eigenen OAuth2-Flow im Browser.

```bash
$ xdg-open "https://login.microsoftonline.com/common/oauth2/v2.0/authorize?
    client_id=9e5f94bc-e8a4-4e73-b8be-63364c29d753&..."
```

Ergebnis:

```
AADSTS700016: Application with identifier '9e5f94bc-...' was not found
in the directory 'Hochschule Düsseldorf'.
```

Die App ist im Tenant **gar nicht registriert**. Thunderbird funktioniert nur, weil es einen alten, noch gültigen Refresh-Token hat, der von vor der Sperrung stammt. Neue Autorisierungen sind blockiert.

### Zweite Erkenntnis

Interessanterweise blockiert unser Hochschul-Tenant nur den **Authorization Code Flow** (Browser-Redirect) für Thunderbirds Client-ID. Der **Device Code Flow** funktioniert dagegen: man gibt einen Code auf `login.microsoft.com/device` ein und meldet sich ganz normal an. Kein Thunderbird nötig.

---

## Lösungsansatz: Device Code Flow + lokaler Proxy

### Architektur

```
Evolution / KMail / Geary / mutt
        |
        | IMAP/SMTP/CalDAV (localhost, kein TLS)
        v
  outlook-linux-proxy
        |
        | XOAUTH2 (IMAP/SMTP) / Bearer Token (EWS)
        v
  outlook.office365.com
```

### Schritt 1: Token holen (ohne Thunderbird)

Das Skript `init-token.py` startet den OAuth2 Device Code Flow:

```bash
$ python3 init-token.py

============================================================
  Öffne im Browser:  https://login.microsoft.com/device
  Gib diesen Code ein:  DMATLB42R
============================================================

Warte auf Anmeldung...
Token gespeichert (gültig für 5198s)
```

Man öffnet die URL, gibt den Code ein, meldet sich mit dem Hochschul-Account an — fertig. Der Proxy verwaltet den Token danach selbständig (automatischer Refresh vor Ablauf, Rotation wird persistiert).

Falls der Device Code Flow im Tenant blockiert wird, gibt es einen Fallback:
```bash
$ python3 init-token.py --thunderbird
```
Das extrahiert den Token aus einer bestehenden Thunderbird-Installation via libnss3.

### Schritt 2: IMAP/SMTP-Proxy

Ein Python-Skript lauscht lokal auf Port 1143 (IMAP) und 1025 (SMTP). Wenn Evolution sich mit `LOGIN user pass` anmeldet, macht der Proxy:

1. SSL-Verbindung zu `outlook.office365.com:993` aufbauen
2. `AUTHENTICATE XOAUTH2` mit dem Access-Token senden
3. Ab jetzt: reiner TCP-Passthrough — jedes Byte wird 1:1 durchgereicht

```
Evolution                    Proxy (localhost)              Microsoft 365
   |                             |                              |
   |-- LOGIN user pass --------->|                              |
   |                             |-- SSL connect --------------->|
   |                             |<-- * OK IMAP ready -----------|
   |                             |-- AUTHENTICATE XOAUTH2 ----->|
   |                             |<-- OK Authenticated ---------|
   |<-- OK LOGIN completed ------|                              |
   |                             |                              |
   |-- SELECT INBOX ------------>|========= passthrough =======>|
   |<-- * 36 EXISTS -------------|<======== passthrough ========|
   |-- FETCH 1:* ... ----------->|========= passthrough =======>|
   |<-- ... --------------------|<======== passthrough ========|
```

Das gleiche Prinzip für SMTP, nur mit STARTTLS upstream.

### Schritt 3: CalDAV-Bridge für den Kalender

Outlook nutzt kein CalDAV, sondern Exchange Web Services (EWS). Also übersetzen wir:

| Evolution schickt (CalDAV) | Bridge macht (EWS) |
|---|---|
| `PROPFIND /calendar/` | `FindItem` mit `CalendarView` |
| `GET /calendar/uid.ics` | iCal aus EWS-Daten bauen |
| `PUT /calendar/uid.ics` mit `PARTSTAT=ACCEPTED` | `AcceptItem` an Meeting-Organisator senden |
| `PUT /calendar/uid.ics` (neues Event) | `CreateItem` |
| `DELETE /calendar/uid.ics` | `DeleteItem` |

Meeting-Einladungen werden als iCal-Events mit `PARTSTAT=NEEDS-ACTION` an Evolution übergeben. Wenn man in Evolution auf "Annehmen" klickt, sendet Evolution ein PUT mit `PARTSTAT=ACCEPTED`, und die Bridge übersetzt das in einen EWS-`AcceptItem`-Aufruf und die Antwort geht an den Organisator.

#### Deduplizierung

Outlook liefert über EWS manchmal denselben Termin mehrfach mit verschiedenen UIDs — besonders bei wiederkehrenden Terminen und Meeting-Einladungen. Die Bridge dedupliziert automatisch: gleicher Betreff + gleiches Startdatum = nur einmal anzeigen.

---

## Was kann dieser Ansatz jetzt?

- **Mail empfangen** (IMAP via `localhost:1143`)
- **Mail senden** (SMTP via `localhost:1025`)
- **Kalender lesen** (~300 Termine, 3 Monate zurück bis 6 Monate voraus)
- **Termine erstellen und löschen**
- **Meeting-Einladungen annehmen, ablehnen, vorläufig zusagen**
- **Automatischer Token-Refresh** (Rotation wird korrekt gehandhabt)
- **systemd-User-Services** mit Autostart beim Login
- **Keine Abhängigkeiten** außer Python 3 Standardbibliothek

---

## Einrichtung

### 1. Repo klonen und installieren

```bash
git clone https://github.com/christiansteier/outlook-linux-proxy.git
cd outlook-linux-proxy
chmod +x install.sh
./install.sh
```

Der Installer fragt nach der Config (`~/.config/outlook-proxy/config.json`), startet den Device Code Flow im Browser und richtet die systemd-Services ein.

### 2. Config anpassen

```json
{
    "tenant_id": "DEINE-AZURE-TENANT-ID",
    "email": "dein.name@deine-hochschule.de",
    "client_id": "9e5f94bc-e8a4-4e73-b8be-63364c29d753",
    "redirect_uri": "https://localhost"
}
```

Die Tenant-ID findet man so:

```bash
python3 -c "
import json, base64
t = json.load(open('$HOME/.config/outlook-proxy/token.json'))
payload = t['access_token'].split('.')[1] + '=='
print('Tenant-ID:', json.loads(base64.b64decode(payload))['tid'])
"
```

### 3. Evolution einrichten

| Einstellung | Wert |
|---|---|
| **IMAP Server** | `localhost`, Port `1143`, keine Verschlüsselung |
| **SMTP Server** | `localhost`, Port `1025`, keine Verschlüsselung |
| **Auth** | Passwort (irgendwas eingeben, wird ignoriert) |
| **CalDAV Kalender** | `http://localhost:1080/calendar/` |

---

## Abläufe

### Normalbetrieb

Der Access-Token ist ~60 Minuten gültig. Wenn er abläuft, holt der Proxy automatisch einen neuen — die Token-Kette erneuert sich endlos, solange der Proxy regelmäßig läuft.

### Meeting-Einladungen

```
  Kollegin schickt                CalDAV-Bridge              Exchange
  Termineinladung                     |                         |
       |                              |                         |
       |  (nächster Kalender-Sync)     |                         |
       |                              |-- FindItem -----------→ |
       |                              |← CalendarItem           |
       |                              |  MyResponseType:        |
       |                              |  NoResponseReceived     |
       |                              |                         |
  Evolution zeigt:                    |                         |
  "Neue Einladung"                    |                         |
  [Annehmen] [Ablehnen]               |                         |
       |                              |                         |
  Du klickst "Annehmen"              |                         |
       |                              |                         |
       |-- PUT event.ics ----------->|                         |
       |   PARTSTAT=ACCEPTED          |                         |
       |                              |-- AcceptItem --------→ |
       |                              |         → Antwort-Mail  |
       |                              |           an Kollegin   |
       |                              |← OK                    |
       |<-- 204 OK ------------------|                         |
```

### Was ist, wenn die Token-Kette reißt?

| Situation | Was passiert | Was tun |
|---|---|---|
| Rechner ein paar Wochen aus | Refresh-Token gilt immerhin 90 Tage | Nichts, Proxy erneuert beim Start |
| Rechner 3+ Monate aus | Refresh-Token abgelaufen | `init-token.py` nochmal ausführen |
| Passwort an der Hochschule geändert | Alle Tokens ungültig | `init-token.py` nochmal ausführen |
| Proxy-Service gecrasht | systemd startet nach 10s neu | Prüfen mit `systemctl --user status ...` |
| Tenant sperrt Client-ID komplett | Auch Device Code Flow geht nicht mehr | Jetzt ist die Campus IT gefragt |

---

## Betrieb

```bash
# Status prüfen
systemctl --user status outlook-mail-proxy outlook-caldav-bridge

# Neustart
systemctl --user restart outlook-mail-proxy outlook-caldav-bridge

# Logs live
journalctl --user -u outlook-mail-proxy -f
journalctl --user -u outlook-caldav-bridge -f

# Token neu holen (nach Passwortänderung)
python3 ~/.local/bin/outlook-proxy-init-token.py
```

---

## Dateien

| Datei | Was sie tut |
|---|---|
| `common.py` | Config laden, Token-Management, Thunderbird-Profil finden |
| `imap-smtp-proxy.py` | IMAP- und SMTP-Proxy mit XOAUTH2 und TCP-Passthrough |
| `caldav-bridge.py` | CalDAV-zu-EWS-Bridge mit Meeting-Einladungen und Deduplizierung |
| `init-token.py` | Device Code Flow oder Token aus Thunderbird extrahieren (via libnss3) |
| `install.sh` | Installer: Config, Services, Token-Init |
| `config.example.json` | Config-Vorlage (keine echten Daten) |

---

## Sicherheit

- Alle Ports nur auf `127.0.0.1` — nicht aus dem Netz erreichbar
- Token-Datei ist `chmod 600`
- Der Refresh-Token gewährt vollen Zugriff auf Postfach und Kalender — entsprechend schützen
- `config.json` und `token.json` stehen im `.gitignore`
- Die Client-ID ist Thunderbirds öffentliche ID (kein Client-Secret, Public-Client-Flow)

---

## Voraussetzungen

- Linux mit systemd
- Python 3.10+
- Ein Microsoft-365-Konto an einer Organisation, die Thunderbirds OAuth2-Client-ID nicht komplett gesperrt hat
- Optional: Thunderbird + `libnss3` (nur für den Fallback `--thunderbird`)

---

## Getestet auf

| Komponente | Version |
|---|---|
| Betriebssystem | Pop!_OS 24.04 LTS |
| Desktop | System76 COSMIC |
| Kernel | 6.18.7 |
| Python | 3.12 |
| Mailclient | GNOME Evolution 3.52 |
| Microsoft 365 | Exchange Online (Hochschul-Tenant mit eingeschränkten App-Registrierungen) |

Sollte auf jedem Linux mit Python 3.10+ und systemd laufen. Als Mailclient geht alles, was IMAP/SMTP spricht (Evolution, KMail, Geary, mutt, etc.) und optional CalDAV für den Kalender.

---

## Einschränkungen

- Wenn der Tenant-Admin Thunderbirds Client-ID oder den Device Code Flow sperrt, funktioniert auch dieser Proxy nicht mehr
- Die CalDAV-Bridge unterstützt keine wiederkehrenden Terminserien mit Ausnahmen und kein komplexes Teilnehmermanagement
- Kontakt-Sync (CardDAV) ist nicht implementiert

---

## Lizenz

MIT
