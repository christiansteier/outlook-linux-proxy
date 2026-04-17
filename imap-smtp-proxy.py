#!/usr/bin/env python3
"""
IMAP + SMTP Proxy für Microsoft 365 Outlook mit OAuth2.
Lauscht lokal und authentifiziert via XOAUTH2 upstream.
Nach dem Login wird die Verbindung als TCP-Passthrough weitergeleitet.
"""

import base64
import os
import signal
import socket
import ssl
import sys
import threading
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import load_config, get_access_token


def read_line(sock):
    buf = b""
    while True:
        ch = sock.recv(1)
        if not ch:
            return None
        buf += ch
        if buf.endswith(b"\r\n"):
            return buf


def pipe(src, dst):
    try:
        while True:
            data = src.recv(8192)
            if not data:
                break
            dst.sendall(data)
    except (OSError, BrokenPipeError):
        pass
    try:
        dst.shutdown(socket.SHUT_WR)
    except OSError:
        pass


# ─── IMAP ───

def handle_imap_client(client_sock, addr, cfg):
    upstream = None
    try:
        client_sock.sendall(b"* OK IMAP4rev1 Outlook-Proxy ready\r\n")

        while True:
            line = read_line(client_sock)
            if not line:
                return

            line_str = line.decode("utf-8", errors="replace").strip()
            parts = line_str.split(" ", 2)
            if len(parts) < 2:
                continue

            tag = parts[0]
            cmd = parts[1].upper()

            if cmd == "CAPABILITY":
                client_sock.sendall(
                    f"* CAPABILITY IMAP4rev1 AUTH=LOGIN AUTH=PLAIN MOVE SPECIAL-USE "
                    f"UIDPLUS NAMESPACE LITERAL+\r\n"
                    f"{tag} OK CAPABILITY completed\r\n".encode()
                )

            elif cmd == "LOGIN":
                rest = parts[2] if len(parts) > 2 else ""
                if rest.startswith('"'):
                    try:
                        end = rest.index('"', 1)
                        user = rest[1:end]
                    except ValueError:
                        user = rest.strip('"').split()[0]
                else:
                    user = rest.split()[0]

                access_token = get_access_token()
                if not access_token:
                    client_sock.sendall(f"{tag} NO Token nicht verfuegbar\r\n".encode())
                    continue

                try:
                    ctx = ssl.create_default_context()
                    raw = socket.create_connection(
                        (cfg["imap_host"], cfg["imap_port"]), timeout=30
                    )
                    upstream = ctx.wrap_socket(raw, server_hostname=cfg["imap_host"])
                except Exception as e:
                    client_sock.sendall(f"{tag} NO Verbindung fehlgeschlagen: {e}\r\n".encode())
                    continue

                greeting = read_line(upstream)
                if not greeting:
                    client_sock.sendall(f"{tag} NO Kein Greeting\r\n".encode())
                    upstream.close()
                    upstream = None
                    continue

                auth_string = f"user={user}\x01auth=Bearer {access_token}\x01\x01"
                auth_b64 = base64.b64encode(auth_string.encode()).decode()
                upstream.sendall(f"A001 AUTHENTICATE XOAUTH2 {auth_b64}\r\n".encode())

                auth_resp = read_line(upstream)
                if auth_resp and b"A001 OK" in auth_resp:
                    print(f"[imap] {user} OK")
                    client_sock.sendall(f"{tag} OK LOGIN completed\r\n".encode())
                    t1 = threading.Thread(target=pipe, args=(client_sock, upstream), daemon=True)
                    t2 = threading.Thread(target=pipe, args=(upstream, client_sock), daemon=True)
                    t1.start()
                    t2.start()
                    t1.join()
                    t2.join()
                    return
                else:
                    err = auth_resp.decode(errors="replace").strip() if auth_resp else "no response"
                    print(f"[imap] Auth fehlgeschlagen: {err}")
                    client_sock.sendall(f"{tag} NO {err}\r\n".encode())
                    upstream.close()
                    upstream = None

            elif cmd == "LOGOUT":
                client_sock.sendall(f"* BYE\r\n{tag} OK LOGOUT completed\r\n".encode())
                return
            else:
                client_sock.sendall(f"{tag} BAD Not authenticated\r\n".encode())

    except (BrokenPipeError, ConnectionResetError, OSError):
        pass
    finally:
        try:
            client_sock.close()
        except:
            pass
        if upstream:
            try:
                upstream.close()
            except:
                pass


# ─── SMTP ───

def handle_smtp_client(client_sock, addr, cfg):
    upstream = None
    try:
        client_sock.sendall(b"220 Outlook-SMTP-Proxy ready\r\n")

        while True:
            line = read_line(client_sock)
            if not line:
                return

            line_str = line.decode("utf-8", errors="replace").strip()
            if not line_str:
                continue
            cmd = line_str.split()[0].upper()

            if cmd == "EHLO":
                client_sock.sendall(
                    b"250-localhost\r\n250-AUTH LOGIN PLAIN\r\n"
                    b"250-8BITMIME\r\n250 SIZE 157286400\r\n"
                )

            elif cmd == "STARTTLS":
                client_sock.sendall(b"502 Not needed on localhost\r\n")

            elif cmd == "AUTH":
                parts = line_str.split()
                user = None

                if len(parts) >= 3 and parts[1].upper() == "PLAIN":
                    decoded = base64.b64decode(parts[2]).decode()
                    fields = decoded.split("\x00")
                    user = fields[1] if len(fields) > 1 else fields[0]
                elif parts[1].upper() == "LOGIN":
                    client_sock.sendall(b"334 VXNlcm5hbWU6\r\n")
                    user_line = read_line(client_sock)
                    if user_line:
                        user = base64.b64decode(user_line.strip()).decode()
                    client_sock.sendall(b"334 UGFzc3dvcmQ6\r\n")
                    read_line(client_sock)

                if not user:
                    client_sock.sendall(b"535 Auth failed\r\n")
                    continue

                access_token = get_access_token()
                if not access_token:
                    client_sock.sendall(b"535 Token not available\r\n")
                    continue

                try:
                    upstream = socket.create_connection(
                        (cfg["smtp_host"], cfg["smtp_port"]), timeout=30
                    )
                    read_line(upstream)  # greeting
                    upstream.sendall(b"EHLO localhost\r\n")
                    while True:
                        r = read_line(upstream)
                        if not r or r[3:4] == b" ":
                            break
                    upstream.sendall(b"STARTTLS\r\n")
                    read_line(upstream)
                    ctx = ssl.create_default_context()
                    upstream = ctx.wrap_socket(upstream, server_hostname=cfg["smtp_host"])
                    upstream.sendall(b"EHLO localhost\r\n")
                    while True:
                        r = read_line(upstream)
                        if not r or r[3:4] == b" ":
                            break

                    auth_string = f"user={user}\x01auth=Bearer {access_token}\x01\x01"
                    auth_b64 = base64.b64encode(auth_string.encode()).decode()
                    upstream.sendall(f"AUTH XOAUTH2 {auth_b64}\r\n".encode())
                    auth_resp = read_line(upstream)

                    if auth_resp and auth_resp.startswith(b"235"):
                        print(f"[smtp] {user} OK")
                        client_sock.sendall(b"235 2.7.0 Authentication successful\r\n")
                        t1 = threading.Thread(target=pipe, args=(client_sock, upstream), daemon=True)
                        t2 = threading.Thread(target=pipe, args=(upstream, client_sock), daemon=True)
                        t1.start()
                        t2.start()
                        t1.join()
                        t2.join()
                        return
                    else:
                        err = auth_resp.decode(errors="replace").strip() if auth_resp else "?"
                        client_sock.sendall(f"535 {err}\r\n".encode())
                        upstream.close()
                        upstream = None

                except Exception as e:
                    client_sock.sendall(f"535 {e}\r\n".encode())
                    if upstream:
                        try:
                            upstream.close()
                        except:
                            pass
                        upstream = None

            elif cmd == "QUIT":
                client_sock.sendall(b"221 Bye\r\n")
                return
            elif cmd in ("NOOP", "RSET"):
                client_sock.sendall(b"250 OK\r\n")
            else:
                client_sock.sendall(b"503 Not authenticated\r\n")

    except (BrokenPipeError, ConnectionResetError, OSError):
        pass
    finally:
        try:
            client_sock.close()
        except:
            pass
        if upstream:
            try:
                upstream.close()
            except:
                pass


def run_server(bind, port, handler, cfg):
    srv = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    srv.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
    srv.bind((bind, port))
    srv.listen(5)
    while True:
        try:
            client, addr = srv.accept()
            t = threading.Thread(target=handler, args=(client, addr, cfg), daemon=True)
            t.start()
        except OSError:
            break


def main():
    cfg = load_config()

    token = get_access_token()
    if not token:
        print("Token ungueltig. Bitte init-token.py ausfuehren.")
        sys.exit(1)

    imap_port = cfg.get("local_imap_port", 1143)
    smtp_port = cfg.get("local_smtp_port", 1025)
    print(f"IMAP: localhost:{imap_port}  SMTP: localhost:{smtp_port}")

    threading.Thread(
        target=run_server, args=("127.0.0.1", imap_port, handle_imap_client, cfg), daemon=True
    ).start()
    threading.Thread(
        target=run_server, args=("127.0.0.1", smtp_port, handle_smtp_client, cfg), daemon=True
    ).start()

    signal.signal(signal.SIGINT, lambda s, f: sys.exit(0))
    signal.signal(signal.SIGTERM, lambda s, f: sys.exit(0))

    print("Proxy laeuft.")
    while True:
        time.sleep(1)


if __name__ == "__main__":
    main()
