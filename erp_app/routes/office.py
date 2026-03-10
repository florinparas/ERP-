"""
Rute pentru integrarea cu suita Microsoft Office.
Export Excel, generare Word, trimitere email Outlook.
"""
from flask import Blueprint, send_file, request, flash, redirect, url_for, render_template
from flask_login import login_required
from erp_app.services.excel_service import (
    export_clients, export_products, export_invoices,
    export_orders, export_employees, export_inventory,
    import_clients, import_products,
)
from erp_app.services.word_service import (
    generate_invoice_doc, generate_order_doc, generate_report_doc,
)
from erp_app.services.email_service import send_invoice_email, send_order_confirmation

office_bp = Blueprint("office", __name__)


# ── Excel Exports ──────────────────────────────────────────

@office_bp.route("/excel/clients")
@login_required
def excel_clients():
    output = export_clients()
    return send_file(output, download_name="clienti.xlsx", as_attachment=True,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@office_bp.route("/excel/products")
@login_required
def excel_products():
    output = export_products()
    return send_file(output, download_name="produse.xlsx", as_attachment=True,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@office_bp.route("/excel/invoices")
@login_required
def excel_invoices():
    output = export_invoices()
    return send_file(output, download_name="facturi.xlsx", as_attachment=True,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@office_bp.route("/excel/orders")
@login_required
def excel_orders():
    output = export_orders()
    return send_file(output, download_name="comenzi.xlsx", as_attachment=True,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@office_bp.route("/excel/employees")
@login_required
def excel_employees():
    output = export_employees()
    return send_file(output, download_name="angajati.xlsx", as_attachment=True,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@office_bp.route("/excel/inventory")
@login_required
def excel_inventory():
    output = export_inventory()
    return send_file(output, download_name="stocuri.xlsx", as_attachment=True,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ── Excel Imports ──────────────────────────────────────────

@office_bp.route("/import/clients", methods=["POST"])
@login_required
def import_clients_route():
    file = request.files.get("file")
    if not file or not file.filename.endswith(".xlsx"):
        flash("Selectati un fisier .xlsx valid!", "error")
        return redirect(request.referrer or url_for("clients.list_clients"))
    imported, errors = import_clients(file)
    flash(f"{imported} clienti importati din Excel!", "success")
    if errors:
        for err in errors[:5]:
            flash(err, "warning")
    return redirect(url_for("clients.list_clients"))


@office_bp.route("/import/products", methods=["POST"])
@login_required
def import_products_route():
    file = request.files.get("file")
    if not file or not file.filename.endswith(".xlsx"):
        flash("Selectati un fisier .xlsx valid!", "error")
        return redirect(request.referrer or url_for("products.list_products"))
    imported, errors = import_products(file)
    flash(f"{imported} produse importate din Excel!", "success")
    if errors:
        for err in errors[:5]:
            flash(err, "warning")
    return redirect(url_for("products.list_products"))


# ── Word Documents ─────────────────────────────────────────

@office_bp.route("/word/invoice/<int:id>")
@login_required
def word_invoice(id):
    output = generate_invoice_doc(id)
    if not output:
        flash("Factura nu a fost gasita!", "error")
        return redirect(url_for("invoices.list_invoices"))
    return send_file(output, download_name=f"factura_{id}.docx", as_attachment=True,
                     mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")


@office_bp.route("/word/order/<int:id>")
@login_required
def word_order(id):
    output = generate_order_doc(id)
    if not output:
        flash("Comanda nu a fost gasita!", "error")
        return redirect(url_for("orders.list_orders"))
    return send_file(output, download_name=f"comanda_{id}.docx", as_attachment=True,
                     mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")


@office_bp.route("/word/report")
@login_required
def word_report():
    output = generate_report_doc()
    return send_file(output, download_name="raport_erp.docx", as_attachment=True,
                     mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")


# ── Email / Outlook ────────────────────────────────────────

@office_bp.route("/email/invoice/<int:id>", methods=["POST"])
@login_required
def email_invoice(id):
    success, message = send_invoice_email(id)
    flash(message, "success" if success else "error")
    return redirect(url_for("invoices.view_invoice", id=id))


@office_bp.route("/email/order/<int:id>", methods=["POST"])
@login_required
def email_order(id):
    success, message = send_order_confirmation(id)
    flash(message, "success" if success else "error")
    return redirect(url_for("orders.view_order", id=id))
