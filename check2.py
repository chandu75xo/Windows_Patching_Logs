import subprocess
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import requests
from bs4 import BeautifulSoup

def get_email_credentials():
    while True:
        sender = input("Enter your UBS email address: ")
        if not sender.endswith('@ubs.com'):
            print("❌ Error: Please enter a valid UBS email address (@ubs.com)")
            continue
        
        receiver = input("Enter recipient's UBS email address: ")
        if not receiver.endswith('@ubs.com'):
            print("❌ Error: Please enter a valid UBS email address (@ubs.com)")
            continue
            
        password = input("Enter your email password: ")
        smtp_server = input("Enter SMTP server address (e.g., smtp.ubs.com): ")
        
        return sender, receiver, password, smtp_server


def check_installed_updates():
    try:
        cmd = 'powershell "Get-WmiObject -Namespace \\"root\\cimv2\\" -Class Win32_QuickFixEngineering | Select-Object -Property CSName, HotFixID, InstalledOn"'
        result = subprocess.run(cmd, capture_output=True, text=True, shell=True)
        return result.stdout
    except Exception as e:
        return f"Failed to fetch updates: {e}"


def get_kb_details(kb_id):
    try:
        kb_number = kb_id.upper().replace("KB", "")
        url = f"https://support.microsoft.com/help/{kb_number}"
        headers = {
            "User-Agent": "Mozilla/5.0"
        }
        response = requests.get(url, headers=headers)
        soup = BeautifulSoup(response.text, 'html.parser')

        title_tag = soup.find('title')
        title = title_tag.get_text().strip() if title_tag else "No title found"

        summary_tag = soup.find('meta', {'name': 'description'})
        summary = summary_tag['content'].strip() if summary_tag else "No summary available"

        return {
            "title": title,
            "summary": summary,
            "link": url
        }
    except Exception as e:
        return {
            "title": "Error fetching details",
            "summary": str(e),
            "link": "#"
        }


def parse_updates(output):
    lines = output.splitlines()
    updates = []
    for line in lines:
        if "KB" in line:
            parts = line.split()
            if len(parts) >= 2:
                kb_id = parts[1]
                kb_info = get_kb_details(kb_id)
                update = {
                    "computer": parts[0],
                    "title": kb_id,
                    "status": "Installed",
                    "summary": kb_info['summary'],
                    "link": kb_info['link']
                }
                updates.append(update)
    return updates or [{
        "computer": "N/A", "title": "No KB updates found", "status": "N/A", "summary": "", "link": "#"
    }]


def generate_html_report(updates):
    table_rows = ""
    for update in updates:
        table_rows += f"""
        <tr>
            <td>{update['computer']}</td>
            <td><a href="{update['link']}" target="_blank">{update['title']}</a></td>
            <td style='color: green;'>{update['status']}</td>
            <td>{update['summary']}</td>
        </tr>
        """

    html_body = f"""
    <html>
    <body>
        <p>Hi Team,</p>
        <p>Here is the latest patch report from our Windows infrastructure:</p>
        <table border="1" cellpadding="6" cellspacing="0" style="border-collapse: collapse; font-family: Arial;">
            <tr style="background-color: #f2f2f2;">
                <th>Computer</th>
                <th>Update</th>
                <th>Status</th>
                <th>Summary</th>
            </tr>
            {table_rows}
        </table>
        <br>
        <p>Kind regards,<br>Automation Script 🤖</p>
    </body>
    </html>
    """
    return html_body


def send_email(subject, html_body, sender, receiver, password, smtp_server):
    msg = MIMEMultipart("alternative")
    msg['Subject'] = subject
    msg['From'] = sender
    msg['To'] = receiver

    html_part = MIMEText(html_body, "html")
    msg.attach(html_part)

    try:
        SMTP_PORT = 587  # Common port for TLS

        with smtplib.SMTP(smtp_server, SMTP_PORT) as server:
            server.starttls()
            server.login(sender, password)
            server.sendmail(sender, receiver, msg.as_string())
        print("✅ Email sent successfully!")
    except Exception as e:
        print("❌ Failed to send email:", e)


if __name__ == "__main__":
    print("📧 Email Configuration")
    print("-" * 50)
    sender, receiver, password, smtp_server = get_email_credentials()
    
    print("\n🔍 Checking installed updates...")
    raw_output = check_installed_updates()
    parsed_updates = parse_updates(raw_output)
    html_report = generate_html_report(parsed_updates)

    print("📤 Sending update report via email...")
    send_email("🖥️ Windows Patch Report", html_report, sender, receiver, password, smtp_server)
