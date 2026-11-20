import imaplib
import email
from email.header import decode_header
import os
import requests

def fetch_recent_emails(limit=15):
    host = os.environ.get("OUTLOOK_IMAP_HOST", "outlook.office365.com")
    user = os.environ["OUTLOOK_EMAIL"]
    password = os.environ["OUTLOOK_APP_PASSWORD"]

    mail = imaplib.IMAP4_SSL(host)
    mail.login(user, password)
    mail.select("INBOX")

    _, messages = mail.search(None, "ALL")
    email_ids = messages[0].split()[-limit:]
    summaries = []

    for eid in reversed(email_ids):
        _, msg_data = mail.fetch(eid, "(RFC822)")
        raw_msg = email.message_from_bytes(msg_data[0][1])
        subject_parts = decode_header(raw_msg.get("Subject", ""))
        subject = "".join(
            part.decode(enc or "utf-8", errors="ignore") if isinstance(part, bytes) else part
            for part, enc in subject_parts
        )
        sender = raw_msg.get("From", "")
        date = raw_msg.get("Date", "")
        summaries.append(f"- [{date}] De: {sender} | Objet: {subject}")

    mail.logout()
    return "\n".join(summaries) if summaries else "Aucun e-mail récent."

def fetch_cpanel_status():
    base_url = os.environ["CPANEL_URL"].rstrip("/")
    user = os.environ["CPANEL_USER"]
    token = os.environ["CPANEL_API_TOKEN"]
    headers = {"Authorization": f"cpanel {user}:{token}"}

    cron_res = requests.get(
        f"{base_url}/execute/Cron/list_cron",
        headers=headers,
        timeout=15
    )
    cron_data = cron_res.json().get("data", []) if cron_res.ok else []
    crons = [
        f"- Cron: {item.get('command', '')} ({item.get('minute')} {item.get('hour')} {item.get('day')} {item.get('month')} {item.get('weekday')})"
        for item in cron_data[:15]
    ]

    return "\n".join(crons) if crons else "Aucune tâche cron récupérée ou accès limité."

def generate_briefing(emails_text, cpanel_text):
    api_key = os.environ["LLM_API_KEY"]
    prompt = f"""
Tu es un assistant personnel. À partir des données suivantes, fais un point clair et concis
des activités à venir, échéances, alertes et actions requises (en français).

=== E-MAILS RÉCENTS (OUTLOOK) ===
{emails_text}

=== ÉTAT CPANEL ===
{cpanel_text}
"""
    # Exemple compatible endpoint OpenAI-like (adaptable selon votre fournisseur gratuit)
    response = requests.post(
        "https://api.openai.com/v1/chat/completions",
        headers={
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
        },
        json={
            "model": "gpt-4o-mini",
            "messages": [{"role": "user", "content": prompt}],
            "temperature": 0.2,
        },
        timeout=30,
    )
    response.raise_for_status()
    return response.json()["choices"][0]["message"]["content"]

def send_telegram(message):
    token = os.environ["TELEGRAM_BOT_TOKEN"]
    chat_id = os.environ["TELEGRAM_CHAT_ID"]
    url = f"https://api.telegram.org/bot{token}/sendMessage"
    # Découpage si > 4000 caractères (limite Telegram = 4096)
    for i in range(0, len(message), 4000):
        requests.post(
            url,
            json={"chat_id": chat_id, "text": message[i:i+4000]},
            timeout=15,
        ).raise_for_status()

if __name__ == "__main__":
    emails_info = fetch_recent_emails()
    cpanel_info = fetch_cpanel_status()
    report = generate_briefing(emails_info, cpanel_info)
    send_telegram(report)