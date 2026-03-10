"""
Serviciu de integrare cu Microsoft Outlook / SMTP.
Trimitere notificari email pentru facturi, comenzi, rapoarte.
"""
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from flask import current_app
from erp_app.models import Invoice, Order


def send_email(to_email, subject, body_html, attachments=None):
    """Trimite email via SMTP (compatibil Outlook/Office 365)."""
    mail_server = current_app.config.get("MAIL_SERVER")
    mail_port = current_app.config.get("MAIL_PORT", 587)
    mail_username = current_app.config.get("MAIL_USERNAME")
    mail_password = current_app.config.get("MAIL_PASSWORD")
    mail_sender = current_app.config.get("MAIL_DEFAULT_SENDER", mail_username)

    if not mail_username or not mail_password:
        return False, "Configurare email lipsa. Setati MAIL_USERNAME si MAIL_PASSWORD."

    msg = MIMEMultipart()
    msg["From"] = mail_sender
    msg["To"] = to_email
    msg["Subject"] = subject

    msg.attach(MIMEText(body_html, "html"))

    if attachments:
        for filename, file_data in attachments:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(file_data.read() if hasattr(file_data, "read") else file_data)
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", f"attachment; filename={filename}")
            msg.attach(part)

    try:
        server = smtplib.SMTP(mail_server, mail_port)
        server.starttls()
        server.login(mail_username, mail_password)
        server.send_message(msg)
        server.quit()
        return True, "Email trimis cu succes!"
    except Exception as e:
        return False, f"Eroare la trimiterea emailului: {str(e)}"


def send_invoice_email(invoice_id):
    """Trimite factura pe email catre client."""
    from erp_app.services.word_service import generate_invoice_doc

    invoice = Invoice.query.get(invoice_id)
    if not invoice or not invoice.client or not invoice.client.email:
        return False, "Factura sau emailul clientului nu exista."

    doc_file = generate_invoice_doc(invoice_id)
    subject = f"Factura {invoice.number} - {invoice.date.strftime('%d.%m.%Y') if invoice.date else ''}"

    body = f"""
    <html>
    <body style="font-family: Arial, sans-serif;">
        <h2>Factura {invoice.number}</h2>
        <p>Stimate/Stimata {invoice.client.name},</p>
        <p>Va transmitem atasat factura <strong>{invoice.number}</strong>
        in valoare de <strong>{invoice.total:.2f} RON</strong>.</p>
        <table style="border-collapse: collapse; margin: 20px 0;">
            <tr><td style="padding: 5px 15px; border: 1px solid #ddd;"><strong>Nr. Factura:</strong></td>
                <td style="padding: 5px 15px; border: 1px solid #ddd;">{invoice.number}</td></tr>
            <tr><td style="padding: 5px 15px; border: 1px solid #ddd;"><strong>Data:</strong></td>
                <td style="padding: 5px 15px; border: 1px solid #ddd;">{invoice.date.strftime('%d.%m.%Y') if invoice.date else '-'}</td></tr>
            <tr><td style="padding: 5px 15px; border: 1px solid #ddd;"><strong>Scadenta:</strong></td>
                <td style="padding: 5px 15px; border: 1px solid #ddd;">{invoice.due_date.strftime('%d.%m.%Y') if invoice.due_date else '-'}</td></tr>
            <tr><td style="padding: 5px 15px; border: 1px solid #ddd;"><strong>Total:</strong></td>
                <td style="padding: 5px 15px; border: 1px solid #ddd;"><strong>{invoice.total:.2f} RON</strong></td></tr>
        </table>
        <p>Cu stima,<br>Echipa ERP</p>
    </body>
    </html>
    """

    attachments = [(f"Factura_{invoice.number}.docx", doc_file)]
    return send_email(invoice.client.email, subject, body, attachments)


def send_order_confirmation(order_id):
    """Trimite confirmare comanda pe email."""
    order = Order.query.get(order_id)
    if not order or not order.client or not order.client.email:
        return False, "Comanda sau emailul clientului nu exista."

    subject = f"Confirmare Comanda {order.number}"
    body = f"""
    <html>
    <body style="font-family: Arial, sans-serif;">
        <h2>Confirmare Comanda {order.number}</h2>
        <p>Stimate/Stimata {order.client.name},</p>
        <p>Comanda dumneavoastra <strong>{order.number}</strong> a fost confirmata.</p>
        <p><strong>Total:</strong> {order.total:.2f} RON</p>
        <p><strong>Data livrare estimata:</strong> {order.delivery_date.strftime('%d.%m.%Y') if order.delivery_date else 'In curs de stabilire'}</p>
        <p>Cu stima,<br>Echipa ERP</p>
    </body>
    </html>
    """
    return send_email(order.client.email, subject, body)
