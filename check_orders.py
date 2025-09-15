#!/usr/bin/env python3
# Reads your published CSV (Google Sheet), checks Gmail via IMAP for each Order Number,
# detects Pending/Shipped/Delivered, extracts tracking, and writes data/status.csv.

import os, re, csv, imaplib, email, sys
from email.header import decode_header
from html import unescape
from urllib.request import urlopen, Request

IMAP_SERVER = "imap.gmail.com"
IMAP_EMAIL = os.environ.get("IMAP_EMAIL")                 # set in GitHub Secrets
IMAP_APP_PASSWORD = os.environ.get("IMAP_APP_PASSWORD")   # set in GitHub Secrets
SHEET_CSV_URL = os.environ.get("SHEET_CSV_URL")
MAILBOX = os.environ.get("MAILBOX", "INBOX")

RE_UPS   = re.compile(r"\b1Z[0-9A-Z]{16}\b", re.I)
RE_USPS  = re.compile(r"\b9\d{21,22}\b")
RE_FEDEX = re.compile(r"\b\d{12,22}\b")
RE_DHL   = re.compile(r"\b\d{10}\b")
RE_TBA   = re.compile(r"\bTBA[0-9A-Z]+\b", re.I)

STATUS_HINTS = {
    "delivered": (" delivered ", " has been delivered", "was delivered", "delivered on"),
    "shipped":   (" shipped ", " on the way ", " out for delivery ", " has shipped "),
    "pending":   (" order confirmed ", " order confirmation ", " thanks for your order ", " placed "),
}

def die(msg): print(msg, file=sys.stderr); sys.exit(1)

def fetch_sheet_rows(url):
    req = Request(url, headers={"User-Agent":"Mozilla/5.0"})
    with urlopen(req, timeout=30) as r:
        content = r.read().decode("utf-8", errors="ignore")
    rows = list(csv.reader(content.splitlines()))
    if not rows: return []
    header = [h.strip().lower() for h in rows[0]]
    data = []
    for r in rows[1:]:
        row = {header[i]: (r[i].strip() if i < len(r) else "") for i in range(len(header))}
        data.append(row)
    return data

def imap_connect():
    if not (IMAP_EMAIL and IMAP_APP_PASSWORD): die("Missing IMAP_EMAIL/IMAP_APP_PASSWORD.")
    imap = imaplib.IMAP4_SSL(IMAP_SERVER)
    imap.login(IMAP_EMAIL, IMAP_APP_PASSWORD)
    imap.select(MAILBOX)
    return imap

def dec(s):
    if not s: return ""
    out = []
    for t, enc in decode_header(s):
        out.append(t.decode(enc or "utf-8", "ignore") if isinstance(t, bytes) else t)
    return "".join(out)

def html_to_text(h): return unescape(re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", h)))
def get_body(msg):
    if msg.is_multipart():
        for p in msg.walk():
            if p.get_content_type()=="text/plain":
                try: return p.get_payload(decode=True).decode(p.get_content_charset() or "utf-8","ignore")
                except: pass
        for p in msg.walk():
            if p.get_content_type()=="text/html":
                try:
                    return html_to_text(p.get_payload(decode=True).decode(p.get_content_charset() or "utf-8","ignore"))
                except: pass
    else:
        try: return msg.get_payload(decode=True).decode(msg.get_content_charset() or "utf-8","ignore")
        except: return ""
    return ""

def search_latest_for_order(imap, order_number, site_hint=""):
    crit = f'(TEXT "{order_number}")'
    if site_hint: crit = f'(TEXT "{order_number}" TEXT "{site_hint}")'
    typ, data = imap.search(None, crit)
    if typ!="OK" or not data or not data[0]: return None
    latest_id = data[0].split()[-1]
    typ, msg_data = imap.fetch(latest_id, "(RFC822)")
    if typ!="OK" or not msg_data or not msg_data[0]: return None
    return email.message_from_bytes(msg_data[0][1])

def detect_tracking(body):
    for name, rx in (("UPS",RE_UPS),("USPS",RE_USPS),("FedEx",RE_FEDEX),("DHL",RE_DHL),("Amazon Logistics",RE_TBA)):
        m = rx.search(body)
        if m: return name, m.group(0)
    return "", ""

def detect_status(subject, body, has_tracking):
    s = f" {subject.lower()} {body.lower()} "
    if any(h in s for h in STATUS_HINTS["delivered"]): return "Delivered"
    if has_tracking or any(h in s for h in STATUS_HINTS["shipped"]): return "Shipped"
    if any(h in s for h in STATUS_HINTS["pending"]): return "Pending"
    if "order confirmation" in s or "placed" in s: return "Pending"
    return "Pending"

def main():
    if not SHEET_CSV_URL: die("Missing SHEET_CSV_URL.")
    rows = fetch_sheet_rows(SHEET_CSV_URL)
    if not rows: die("No rows from sheet CSV.")

    imap = imap_connect()
    os.makedirs("data", exist_ok=True)
    with open("data/status.csv","w",newline="",encoding="utf-8") as f:
        w = csv.writer(f); w.writerow(["order_number","site","status","tracking","carrier","email_date","subject"])
        for r in rows:
            order = (r.get("order number") or r.get("order_number") or "").strip()
            site  = (r.get("site") or "").strip()
            if not order: continue
            msg = search_latest_for_order(imap, order, site)
            if not msg:
                w.writerow([order, site, "Pending","","","", ""]); continue
            subject = dec(msg.get("Subject") or ""); date_hdr = dec(msg.get("Date") or ""); body = get_body(msg)
            carrier, tracking = detect_tracking(body)
            status = detect_status(subject, body, bool(tracking))
            w.writerow([order, site, status, tracking, carrier, date_hdr, subject])
    imap.logout()
    print("Wrote data/status.csv")

if __name__=="__main__": main()

