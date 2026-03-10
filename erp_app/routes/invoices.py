from datetime import datetime
from flask import Blueprint, render_template, request, redirect, url_for, flash, abort
from flask_login import login_required
from erp_app import db
from erp_app.models.invoice import Invoice, InvoiceItem
from erp_app.models.client import Client
from erp_app.models.product import Product

invoices_bp = Blueprint("invoices", __name__)


def generate_invoice_number():
    last = Invoice.query.order_by(Invoice.id.desc()).first()
    next_num = (last.id + 1) if last else 1
    return f"FACT-{datetime.now().year}-{next_num:05d}"


@invoices_bp.route("/")
@login_required
def list_invoices():
    page = request.args.get("page", 1, type=int)
    status = request.args.get("status", "")
    query = Invoice.query
    if status:
        query = query.filter_by(status=status)
    invoices = query.order_by(Invoice.date.desc()).paginate(page=page, per_page=20)
    return render_template("pages/invoices.html", invoices=invoices, selected_status=status)


@invoices_bp.route("/add", methods=["GET", "POST"])
@login_required
def add_invoice():
    if request.method == "POST":
        invoice = Invoice(
            number=generate_invoice_number(),
            client_id=int(request.form["client_id"]),
            date=datetime.strptime(request.form["date"], "%Y-%m-%d").date() if request.form.get("date") else None,
            due_date=datetime.strptime(request.form["due_date"], "%Y-%m-%d").date() if request.form.get("due_date") else None,
            notes=request.form.get("notes"),
        )
        descriptions = request.form.getlist("item_description[]")
        quantities = request.form.getlist("item_quantity[]")
        prices = request.form.getlist("item_price[]")
        vat_rates = request.form.getlist("item_vat[]")

        for i in range(len(descriptions)):
            if descriptions[i].strip():
                item = InvoiceItem(
                    description=descriptions[i],
                    quantity=float(quantities[i]),
                    unit_price=float(prices[i]),
                    vat_rate=float(vat_rates[i]) if i < len(vat_rates) else 19.0,
                )
                item.calculate()
                invoice.items.append(item)

        invoice.recalculate()
        db.session.add(invoice)
        db.session.commit()
        flash("Factura creata cu succes!", "success")
        return redirect(url_for("invoices.list_invoices"))

    clients = Client.query.order_by(Client.name).all()
    products = Product.query.filter_by(is_active=True).order_by(Product.name).all()
    return render_template(
        "pages/invoice_form.html",
        invoice=None,
        clients=clients,
        products=products,
        invoice_number=generate_invoice_number(),
    )


@invoices_bp.route("/<int:id>")
@login_required
def view_invoice(id):
    invoice = db.session.get(Invoice, id) or abort(404)
    return render_template("pages/invoice_view.html", invoice=invoice)


@invoices_bp.route("/<int:id>/status/<status>", methods=["POST"])
@login_required
def update_status(id, status):
    invoice = db.session.get(Invoice, id) or abort(404)
    if status in ("draft", "sent", "paid", "overdue", "cancelled"):
        invoice.status = status
        db.session.commit()
        flash(f"Status actualizat: {status}", "success")
    return redirect(url_for("invoices.view_invoice", id=id))


@invoices_bp.route("/<int:id>/delete", methods=["POST"])
@login_required
def delete_invoice(id):
    invoice = db.session.get(Invoice, id) or abort(404)
    db.session.delete(invoice)
    db.session.commit()
    flash("Factura stearsa!", "success")
    return redirect(url_for("invoices.list_invoices"))
