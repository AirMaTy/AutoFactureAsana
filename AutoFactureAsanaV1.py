# Version finale d'AutoFacture Asana avec gestion robuste des tâches Asana

import imaplib
import email
import re
import os
from email.header import decode_header
from bs4 import BeautifulSoup
from datetime import datetime
import asana
import tkinter as tk
from tkinter import messagebox, scrolledtext
from dotenv import load_dotenv

# Chargement des variables d'environnement
load_dotenv()

IMAP_SERVER = os.getenv("IMAP_SERVER")
IMAP_PORT = int(os.getenv("IMAP_PORT", 993))
EMAIL_ACCOUNT = os.getenv("IMAP_USERNAME")
EMAIL_PASSWORD = os.getenv("IMAP_PASSWORD")
ASANA_TOKEN = os.getenv("ASANA_API_TOKEN")
PROJECT_GID = os.getenv("PROJECT_GID")
CUSTOM_FIELD_PAIMENT = "1205959178241727"
ENUM_OPTION_PAYE = "1205959178241728"
MAIL_FOLDER = "INBOX/02-STRIPE"

invoice_pattern = re.compile(r"^Votre facture\s+(F\d{8}-\d+)\s+a été payée", re.IGNORECASE)
client_pattern = re.compile(r"Client\s*:\s*(.+)", re.IGNORECASE)

# UI Setup
root = tk.Tk()
root.title("AutoFacture Asana")
root.geometry("700x500")
root.configure(bg="#2e2e2e")

frame = tk.Frame(root, bg="#2e2e2e")
frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

output_text = scrolledtext.ScrolledText(frame, width=90, height=25, bg="#3e3e3e", fg="white", insertbackground="white")
output_text.pack(fill=tk.BOTH, expand=True)

output_text.tag_configure("success", foreground="green")
output_text.tag_configure("failure", foreground="red")
output_text.tag_configure("info", foreground="cyan")

def log(msg, tag=None):
    output_text.insert(tk.END, msg + "\n", tag)
    output_text.see(tk.END)

def decode_mime_words(s):
    return ''.join([
        frag.decode(charset or 'utf-8') if isinstance(frag, bytes) else frag
        for frag, charset in decode_header(s)
    ])

def get_email_content(msg):
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_type() == "text/plain" and "attachment" not in str(part.get("Content-Disposition")):
                return part.get_payload(decode=True).decode(part.get_content_charset() or "utf-8", errors="ignore")
        for part in msg.walk():
            if part.get_content_type() == "text/html" and "attachment" not in str(part.get("Content-Disposition")):
                html = part.get_payload(decode=True).decode(part.get_content_charset() or "utf-8", errors="ignore")
                return BeautifulSoup(html, "html.parser").get_text("\n").strip()
    else:
        body = msg.get_payload(decode=True).decode(msg.get_content_charset() or "utf-8", errors="ignore")
        return BeautifulSoup(body, "html.parser").get_text("\n").strip() if msg.get_content_type() == "text/html" else body
    return ""

def update_task_asana(client, invoice_map, invoice_number, client_name):
    task_gid = invoice_map.get(invoice_number)
    if not task_gid:
        log(f"❌ Tâche non trouvée pour {invoice_number}", "failure")
        return
    client.tasks.update_task(task_gid, {
        "custom_fields": {
            CUSTOM_FIELD_PAIMENT: ENUM_OPTION_PAYE
        }
    })
    log(f"✅ Facture {invoice_number} mise à jour pour {client_name}", "success")

def scan_emails():
    client = asana.Client.access_token(ASANA_TOKEN)
    client.headers["Asana-Disable"] = "new_goal_memberships"
    invoice_map = {}
    log("🔄 Chargement des tâches depuis Asana...", "info")
    try:
        tasks = client.tasks.get_tasks_for_project(PROJECT_GID, opt_fields="name,gid")
        for task in tasks:
            name = task.get("name", "")
            matches = re.findall(r"(F\d{8}-\d+)", name)
            for match in matches:
                invoice_map[match] = task["gid"]

    except Exception as e:
        messagebox.showerror("Erreur Asana", f"Erreur lors de la récupération des tâches :\n{e}")
        return
    log(f"✅ {len(invoice_map)} facture(s) prêtes à être mises à jour.", "success")

    try:
        mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
        mail.login(EMAIL_ACCOUNT, EMAIL_PASSWORD)
        mail.select(MAIL_FOLDER)
        status, data = mail.search(None, "UNSEEN")
        email_ids = data[0].split() if status == "OK" else []
    except Exception as e:
        messagebox.showerror("Erreur IMAP", str(e))
        return

    for eid in email_ids:
        status, data = mail.fetch(eid, "(RFC822)")
        if status != "OK":
            continue
        msg = email.message_from_bytes(data[0][1])
        subject = decode_mime_words(msg.get("Subject", ""))
        if not subject.lower().startswith("votre facture"):
            continue
        invoice_match = invoice_pattern.search(subject)
        if not invoice_match:
            continue
        invoice_number = invoice_match.group(1)
        body = get_email_content(msg)
        client_match = client_pattern.search(body)
        client_name = client_match.group(1).strip() if client_match else "Inconnu"
        update_task_asana(client, invoice_map, invoice_number, client_name)

    mail.logout()
    log("\n🎉 Traitement terminé.", "success")

tk.Button(frame, text="Scanner les e-mails", bg="#4e4e4e", fg="white", font=("Segoe UI", 10, "bold"),
          command=scan_emails).pack(pady=10)

root.mainloop()