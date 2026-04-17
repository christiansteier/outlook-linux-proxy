#!/usr/bin/env python3
"""
CalDAV-zu-EWS Bridge für Microsoft 365 Outlook.
Lokaler CalDAV-Server, der Kalender-Anfragen an Exchange Web Services
weiterleitet (mit OAuth2-Token). Unterstützt Meeting-Einladungen
(Annehmen/Ablehnen/Vorläufig).
"""

import hashlib
import json
import os
import re
import signal
import sys
import time
import urllib.request
from datetime import datetime, timedelta, timezone
from http.server import HTTPServer, BaseHTTPRequestHandler
from socketserver import ThreadingMixIn
from xml.etree import ElementTree as ET

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import load_config, get_access_token

EWS_URL = "https://outlook.office365.com/EWS/Exchange.asmx"

NS = {
    "s": "http://schemas.xmlsoap.org/soap/envelope/",
    "t": "http://schemas.microsoft.com/exchange/services/2006/types",
    "m": "http://schemas.microsoft.com/exchange/services/2006/messages",
}


def ews_request(soap_body):
    token = get_access_token()
    if not token:
        return None

    soap = f'''<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages">
  <soap:Header>
    <t:RequestServerVersion Version="Exchange2016"/>
  </soap:Header>
  <soap:Body>
    {soap_body}
  </soap:Body>
</soap:Envelope>'''

    req = urllib.request.Request(EWS_URL, data=soap.encode("utf-8"), method="POST")
    req.add_header("Authorization", f"Bearer {token}")
    req.add_header("Content-Type", "text/xml; charset=utf-8")

    try:
        resp = urllib.request.urlopen(req, timeout=30)
        return ET.fromstring(resp.read())
    except Exception as e:
        print(f"[caldav] EWS-Fehler: {e}", file=sys.stderr)
        return None


def get_calendar_items(start_date=None, end_date=None):
    if not start_date:
        # Standard: 3 Monate zurück bis 6 Monate voraus
        from datetime import datetime, timedelta
        now = datetime.utcnow()
        start_date = (now - timedelta(days=90)).strftime("%Y-%m-%dT00:00:00Z")
    if not end_date:
        from datetime import datetime, timedelta
        now = datetime.utcnow()
        end_date = (now + timedelta(days=180)).strftime("%Y-%m-%dT23:59:59Z")

    body = f'''<m:FindItem Traversal="Shallow">
      <m:ItemShape>
        <t:BaseShape>AllProperties</t:BaseShape>
        <t:AdditionalProperties>
          <t:FieldURI FieldURI="calendar:UID"/>
          <t:FieldURI FieldURI="calendar:CalendarItemType"/>
          <t:ExtendedFieldURI DistinguishedPropertySetId="PublicStrings"
                              PropertyName="DAV:uid" PropertyType="String"/>
        </t:AdditionalProperties>
      </m:ItemShape>
      <m:CalendarView StartDate="{start_date}" EndDate="{end_date}" MaxEntriesReturned="500"/>
      <m:ParentFolderIds>
        <t:DistinguishedFolderId Id="calendar"/>
      </m:ParentFolderIds>
    </m:FindItem>'''

    root = ews_request(body)
    if root is None:
        return []

    items = []
    for item in root.findall(".//t:CalendarItem", NS):
        items.append(parse_calendar_item(item))
    return items


def get_calendar_item_by_id(item_id):
    body = f'''<m:GetItem>
      <m:ItemShape>
        <t:BaseShape>AllProperties</t:BaseShape>
        <t:BodyType>Text</t:BodyType>
        <t:AdditionalProperties>
          <t:FieldURI FieldURI="calendar:UID"/>
        </t:AdditionalProperties>
      </m:ItemShape>
      <m:ItemIds>
        <t:ItemId Id="{item_id}"/>
      </m:ItemIds>
    </m:GetItem>'''

    root = ews_request(body)
    if root is None:
        return None

    item = root.find(".//t:CalendarItem", NS)
    if item is not None:
        return parse_calendar_item(item)
    return None


def parse_calendar_item(item):
    def text(path):
        el = item.find(path, NS)
        return el.text if el is not None and el.text else ""

    item_id_el = item.find("t:ItemId", NS)
    item_id = item_id_el.get("Id", "") if item_id_el is not None else ""
    change_key = item_id_el.get("ChangeKey", "") if item_id_el is not None else ""

    ews_uid = text("t:UID")
    cal_type = text("t:CalendarItemType")  # Single, Occurrence, Exception
    start_date = text("t:Start")

    # ItemId ist IMMER eindeutig pro Event/Occurrence — als stabile UID nutzen
    # EWS UID ist bei Occurrences oft gleich, ItemId nie
    import hashlib
    uid = hashlib.sha256(item_id.encode()).hexdigest()[:40]

    return {
        "item_id": item_id,
        "change_key": change_key,
        "uid": uid,
        "base_uid": ews_uid,
        "cal_type": cal_type,
        "subject": text("t:Subject"),
        "start": start_date,
        "end": text("t:End"),
        "location": text("t:Location"),
        "body": text("t:Body"),
        "organizer": text(".//t:Organizer/t:Mailbox/t:EmailAddress"),
        "is_all_day": text("t:IsAllDayEvent") == "true",
        "sensitivity": text("t:Sensitivity"),
        "status": text("t:LegacyFreeBusyStatus"),
        "reminder": text("t:ReminderMinutesBeforeStart"),
        "my_response": text("t:MyResponseType"),
        "organizer_email": text(".//t:Organizer/t:Mailbox/t:EmailAddress"),
        "organizer_name": text(".//t:Organizer/t:Mailbox/t:Name"),
    }


def item_to_ical(item):
    uid = item["uid"]
    summary = item["subject"]
    dtstart = format_ical_date(item["start"], item["is_all_day"])
    dtend = format_ical_date(item["end"], item["is_all_day"])
    location = item["location"]
    description = item["body"].replace("\n", "\\n") if item["body"] else ""

    lines = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "PRODID:-//HSD-Proxy//CalDAV Bridge//DE",
        "BEGIN:VEVENT",
        f"UID:{uid}",
        f"SUMMARY:{summary}",
    ]

    if item["is_all_day"]:
        lines.append(f"DTSTART;VALUE=DATE:{dtstart}")
        lines.append(f"DTEND;VALUE=DATE:{dtend}")
    else:
        lines.append(f"DTSTART:{dtstart}")
        lines.append(f"DTEND:{dtend}")

    if location:
        lines.append(f"LOCATION:{location}")
    if description:
        lines.append(f"DESCRIPTION:{description}")
    if item["reminder"]:
        lines.extend([
            "BEGIN:VALARM",
            "ACTION:DISPLAY",
            f"TRIGGER:-PT{item['reminder']}M",
            "DESCRIPTION:Reminder",
            "END:VALARM",
        ])

    status_map = {"Free": "TRANSPARENT", "Busy": "OPAQUE",
                  "Tentative": "TENTATIVE", "OOF": "OPAQUE"}
    transp = status_map.get(item["status"], "OPAQUE")
    lines.append(f"TRANSP:{transp}")

    # Meeting-Einladungen: Organizer + eigener Status
    if item.get("my_response") and item["my_response"] != "Organizer":
        partstat_map = {
            "Accept": "ACCEPTED",
            "Tentative": "TENTATIVE",
            "Decline": "DECLINED",
            "NoResponseReceived": "NEEDS-ACTION",
            "Unknown": "NEEDS-ACTION",
        }
        partstat = partstat_map.get(item["my_response"], "NEEDS-ACTION")

        org_email = item.get("organizer_email", "")
        org_name = item.get("organizer_name", "")
        if org_email:
            if "@" in org_email:
                lines.append(f"ORGANIZER;CN={org_name}:mailto:{org_email}")
            else:
                lines.append(f"ORGANIZER;CN={org_name}:mailto:unknown@unknown")

        lines.append(
            f"ATTENDEE;PARTSTAT={partstat};ROLE=REQ-PARTICIPANT;"
            f"CN={load_config()['email']}:mailto:{load_config()['email']}"
        )
        lines.append(f"METHOD:REQUEST")

    lines.extend(["END:VEVENT", "END:VCALENDAR"])
    return "\r\n".join(lines)


def format_ical_date(datestr, all_day=False):
    if not datestr:
        return ""
    # Parse ISO format
    dt = datestr.replace("Z", "+00:00")
    try:
        d = datetime.fromisoformat(dt)
        if all_day:
            return d.strftime("%Y%m%d")
        return d.strftime("%Y%m%dT%H%M%SZ")
    except:
        return datestr.replace("-", "").replace(":", "").replace(".", "")[:15] + "Z"


def parse_ical_event(ical_text):
    """Einfacher iCal-Parser für Evolution-Events."""
    props = {}
    for line in ical_text.splitlines():
        if ":" not in line:
            continue
        key, _, val = line.partition(":")
        key = key.split(";")[0].strip().upper()
        val = val.strip()
        if key in ("SUMMARY", "LOCATION", "DESCRIPTION", "DTSTART", "DTEND", "UID", "TRANSP"):
            props[key] = val
    return props


def ical_to_ews_datetime(ical_dt):
    """Konvertiert iCal-Datum zu ISO-Format für EWS."""
    ical_dt = ical_dt.strip()
    if len(ical_dt) == 8:  # All day: 20260417
        return f"{ical_dt[:4]}-{ical_dt[4:6]}-{ical_dt[6:8]}T00:00:00Z"
    # 20260417T100000Z
    clean = ical_dt.replace("Z", "")
    return f"{clean[:4]}-{clean[4:6]}-{clean[6:8]}T{clean[9:11]}:{clean[11:13]}:{clean[13:15]}Z"


def create_calendar_item(props):
    subject = props.get("SUMMARY", "Ohne Titel")
    start = ical_to_ews_datetime(props.get("DTSTART", ""))
    end = ical_to_ews_datetime(props.get("DTEND", start))
    location = props.get("LOCATION", "")
    body = props.get("DESCRIPTION", "").replace("\\n", "\n")

    body_xml = f"<t:Body BodyType='Text'>{body}</t:Body>" if body else ""
    location_xml = f"<t:Location>{location}</t:Location>" if location else ""

    soap = f'''<m:CreateItem SendMeetingInvitations="SendToNone">
      <m:Items>
        <t:CalendarItem>
          <t:Subject>{subject}</t:Subject>
          {body_xml}
          {location_xml}
          <t:Start>{start}</t:Start>
          <t:End>{end}</t:End>
        </t:CalendarItem>
      </m:Items>
    </m:CreateItem>'''

    root = ews_request(soap)
    if root is None:
        return False
    resp_class = root.find(".//m:CreateItemResponseMessage", NS)
    if resp_class is not None:
        return resp_class.get("ResponseClass") == "Success"
    return False


def respond_to_meeting(item_id, change_key, response_type):
    """Antwortet auf eine Meeting-Einladung: Accept, Tentative, Decline."""
    response_map = {
        "ACCEPTED": "AcceptItem",
        "TENTATIVE": "TentativelyAcceptItem",
        "DECLINED": "DeclineItem",
    }
    ews_action = response_map.get(response_type)
    if not ews_action:
        print(f"[caldav] Unbekannter Response-Typ: {response_type}", file=sys.stderr)
        return False

    soap = f'''<m:CreateItem MessageDisposition="SendAndSaveCopy">
      <m:Items>
        <t:{ews_action}>
          <t:ReferenceItemId Id="{item_id}" ChangeKey="{change_key}"/>
        </t:{ews_action}>
      </m:Items>
    </m:CreateItem>'''

    root = ews_request(soap)
    if root is None:
        return False
    resp_msg = root.find(".//m:CreateItemResponseMessage", NS)
    if resp_msg is not None and resp_msg.get("ResponseClass") == "Success":
        print(f"[caldav] Meeting-Antwort gesendet: {ews_action}", file=sys.stderr)
        return True
    print(f"[caldav] Meeting-Antwort fehlgeschlagen", file=sys.stderr)
    return False


def delete_calendar_item(item_id, change_key):
    soap = f'''<m:DeleteItem DeleteType="MoveToDeletedItems" SendMeetingCancellations="SendToNone">
      <m:ItemIds>
        <t:ItemId Id="{item_id}" ChangeKey="{change_key}"/>
      </m:ItemIds>
    </m:DeleteItem>'''

    root = ews_request(soap)
    return root is not None


# UID -> ItemId Mapping (Cache)
_uid_cache = {}
_ctag = "initial"
_last_refresh = 0


def refresh_uid_cache(force=False):
    global _uid_cache, _ctag, _last_refresh
    import hashlib

    # Nicht öfter als alle 30 Sekunden refreshen (verhindert Spam bei Evolution-Sync)
    if not force and time.time() - _last_refresh < 30 and _uid_cache:
        return list(_uid_cache.values())

    items = get_calendar_items()

    # Deduplizierung: Outlook liefert manchmal denselben Termin mehrfach
    # (verschiedene UIDs, gleicher Inhalt). Auch escaped Kommas normalisieren.
    def normalize_date(d):
        return d[:10].replace("-", "") if d else ""

    seen = {}
    deduped = []
    for item in items:
        key = (
            item["subject"].lower().strip().replace("\\,", ",").replace("\\;", ";"),
            normalize_date(item["start"]),
        )
        if key not in seen:
            seen[key] = item
            deduped.append(item)
    if len(deduped) < len(items):
        print(f"[caldav] {len(items) - len(deduped)} Duplikate aus Outlook entfernt")
    items = deduped

    _uid_cache = {item["uid"]: item for item in items}
    _last_refresh = time.time()

    # ctag nur ändern wenn sich Inhalte geändert haben
    content_hash = hashlib.md5(
        "".join(f"{i['uid']}{i['change_key']}" for i in items).encode()
    ).hexdigest()[:12]
    _ctag = f"ctag-{content_hash}"

    return items


class CalDAVHandler(BaseHTTPRequestHandler):
    def log_message(self, format, *args):
        print(f"[caldav] {args[0]}", file=sys.stderr)

    def send_xml(self, status, body, content_type="application/xml; charset=utf-8"):
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body.encode("utf-8"))))
        self.send_header("DAV", "1, calendar-access")
        self.end_headers()
        self.wfile.write(body.encode("utf-8"))

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header("Allow", "OPTIONS, GET, PUT, DELETE, PROPFIND, REPORT")
        self.send_header("DAV", "1, calendar-access")
        self.end_headers()

    def do_PROPFIND(self):
        path = self.path.rstrip("/")
        depth = self.headers.get("Depth", "1")

        if path in ("", "/", f"/calendar", f"/calendar/"):
            if depth == "0":
                body = self._propfind_calendar_collection()
            else:
                body = self._propfind_calendar_with_items()
            self.send_xml(207, body)

        elif path == "/.well-known/caldav":
            self.send_response(301)
            self.send_header("Location", "/calendar/")
            self.end_headers()

        else:
            # Einzelnes Event
            uid = path.split("/")[-1].replace(".ics", "")
            if uid in _uid_cache:
                item = _uid_cache[uid]
                etag = f'"{item["change_key"][:20]}"'
                body = f'''<?xml version="1.0" encoding="utf-8"?>
<D:multistatus xmlns:D="DAV:" xmlns:C="urn:ietf:params:xml:ns:caldav">
  <D:response>
    <D:href>{self.path}</D:href>
    <D:propstat>
      <D:prop>
        <D:getetag>{etag}</D:getetag>
        <D:getcontenttype>text/calendar; charset=utf-8</D:getcontenttype>
      </D:prop>
      <D:status>HTTP/1.1 200 OK</D:status>
    </D:propstat>
  </D:response>
</D:multistatus>'''
                self.send_xml(207, body)
            else:
                self.send_error(404)

    def _propfind_calendar_collection(self):
        return f'''<?xml version="1.0" encoding="utf-8"?>
<D:multistatus xmlns:D="DAV:" xmlns:C="urn:ietf:params:xml:ns:caldav"
               xmlns:CS="http://calendarserver.org/ns/">
  <D:response>
    <D:href>/calendar/</D:href>
    <D:propstat>
      <D:prop>
        <D:resourcetype>
          <D:collection/>
          <C:calendar/>
        </D:resourcetype>
        <D:displayname>HSD Kalender</D:displayname>
        <C:supported-calendar-component-set>
          <C:comp name="VEVENT"/>
        </C:supported-calendar-component-set>
        <CS:getctag>{_ctag}</CS:getctag>
      </D:prop>
      <D:status>HTTP/1.1 200 OK</D:status>
    </D:propstat>
  </D:response>
</D:multistatus>'''

    def _propfind_calendar_with_items(self):
        items = refresh_uid_cache()

        responses = [f'''  <D:response>
    <D:href>/calendar/</D:href>
    <D:propstat>
      <D:prop>
        <D:resourcetype><D:collection/><C:calendar/></D:resourcetype>
        <D:displayname>HSD Kalender</D:displayname>
      </D:prop>
      <D:status>HTTP/1.1 200 OK</D:status>
    </D:propstat>
  </D:response>''']

        for item in items:
            etag = f'"{item["change_key"][:20]}"'
            responses.append(f'''  <D:response>
    <D:href>/calendar/{item["uid"]}.ics</D:href>
    <D:propstat>
      <D:prop>
        <D:getetag>{etag}</D:getetag>
        <D:getcontenttype>text/calendar; charset=utf-8</D:getcontenttype>
      </D:prop>
      <D:status>HTTP/1.1 200 OK</D:status>
    </D:propstat>
  </D:response>''')

        return f'''<?xml version="1.0" encoding="utf-8"?>
<D:multistatus xmlns:D="DAV:" xmlns:C="urn:ietf:params:xml:ns:caldav">
{"".join(responses)}
</D:multistatus>'''

    def do_REPORT(self):
        content_length = int(self.headers.get("Content-Length", 0))
        body = self.rfile.read(content_length).decode("utf-8") if content_length else ""

        items = refresh_uid_cache()

        responses = []
        for item in items:
            ical = item_to_ical(item)
            etag = f'"{item["change_key"][:20]}"'
            ical_escaped = ical.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            responses.append(f'''  <D:response>
    <D:href>/calendar/{item["uid"]}.ics</D:href>
    <D:propstat>
      <D:prop>
        <D:getetag>{etag}</D:getetag>
        <C:calendar-data>{ical_escaped}</C:calendar-data>
      </D:prop>
      <D:status>HTTP/1.1 200 OK</D:status>
    </D:propstat>
  </D:response>''')

        result = f'''<?xml version="1.0" encoding="utf-8"?>
<D:multistatus xmlns:D="DAV:" xmlns:C="urn:ietf:params:xml:ns:caldav">
{"".join(responses)}
</D:multistatus>'''
        self.send_xml(207, result)

    def do_GET(self):
        path = self.path.rstrip("/")
        uid = path.split("/")[-1].replace(".ics", "")

        if not _uid_cache:
            refresh_uid_cache()

        if uid in _uid_cache:
            item = _uid_cache[uid]
            ical = item_to_ical(item)
            self.send_response(200)
            self.send_header("Content-Type", "text/calendar; charset=utf-8")
            self.send_header("ETag", f'"{item["change_key"][:20]}"')
            data = ical.encode("utf-8")
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
        else:
            self.send_error(404)

    def do_PUT(self):
        content_length = int(self.headers.get("Content-Length", 0))
        ical_data = self.rfile.read(content_length).decode("utf-8") if content_length else ""

        props = parse_ical_event(ical_data)
        if not props:
            self.send_error(400)
            return

        # Check ob es ein Meeting-Response ist (PARTSTAT-Änderung)
        path = self.path.rstrip("/")
        uid = path.split("/")[-1].replace(".ics", "")

        if not _uid_cache:
            refresh_uid_cache()

        if uid in _uid_cache:
            existing = _uid_cache[uid]
            # PARTSTAT aus dem iCal extrahieren
            partstat = None
            for line in ical_data.splitlines():
                if "PARTSTAT=" in line and load_config()["email"].split("@")[0] in line.lower():
                    import re
                    m = re.search(r"PARTSTAT=(\w+)", line)
                    if m:
                        partstat = m.group(1)
                        break

            if partstat and existing.get("my_response") != "Organizer":
                if respond_to_meeting(existing["item_id"], existing["change_key"], partstat):
                    refresh_uid_cache()
                    self.send_response(204)
                    self.end_headers()
                    return
                else:
                    self.send_error(500)
                    return

        # Normaler neuer Termin
        if create_calendar_item(props):
            refresh_uid_cache()
            self.send_response(201)
            self.end_headers()
        else:
            self.send_error(500)

    def do_DELETE(self):
        path = self.path.rstrip("/")
        uid = path.split("/")[-1].replace(".ics", "")

        if not _uid_cache:
            refresh_uid_cache()

        if uid in _uid_cache:
            item = _uid_cache[uid]
            if delete_calendar_item(item["item_id"], item["change_key"]):
                del _uid_cache[uid]
                self.send_response(204)
                self.end_headers()
            else:
                self.send_error(500)
        else:
            self.send_error(404)


def main():
    cfg = load_config()
    port = cfg.get("local_caldav_port", 1080)

    token = get_access_token()
    if not token:
        print("Token ungueltig. Bitte init-token.py ausfuehren.")
        sys.exit(1)

    items = refresh_uid_cache()
    print(f"[caldav] {len(items)} Termine geladen")

    class ReusableHTTPServer(ThreadingMixIn, HTTPServer):
        allow_reuse_address = True
        daemon_threads = True

    server = ReusableHTTPServer(("127.0.0.1", port), CalDAVHandler)

    signal.signal(signal.SIGINT, lambda s, f: (server.shutdown(), sys.exit(0)))
    signal.signal(signal.SIGTERM, lambda s, f: (server.shutdown(), sys.exit(0)))

    print(f"[caldav] http://localhost:{port}/calendar/")
    server.serve_forever()


if __name__ == "__main__":
    main()
