from datetime import datetime
from flask import Blueprint, render_template, request, redirect, url_for, flash, abort
from flask_login import login_required
from erp_app import db
from erp_app.models.order import Order, OrderItem
from erp_app.models.client import Client
from erp_app.models.product import Product

orders_bp = Blueprint("orders", __name__)


def generate_order_number():
    last = Order.query.order_by(Order.id.desc()).first()
    next_num = (last.id + 1) if last else 1
    return f"CMD-{datetime.now().year}-{next_num:05d}"


@orders_bp.route("/")
@login_required
def list_orders():
    page = request.args.get("page", 1, type=int)
    status = request.args.get("status", "")
    query = Order.query
    if status:
        query = query.filter_by(status=status)
    orders = query.order_by(Order.date.desc()).paginate(page=page, per_page=20)
    return render_template("pages/orders.html", orders=orders, selected_status=status)


@orders_bp.route("/add", methods=["GET", "POST"])
@login_required
def add_order():
    if request.method == "POST":
        order = Order(
            number=generate_order_number(),
            client_id=int(request.form["client_id"]),
            date=datetime.strptime(request.form["date"], "%Y-%m-%d").date() if request.form.get("date") else None,
            delivery_date=datetime.strptime(request.form["delivery_date"], "%Y-%m-%d").date() if request.form.get("delivery_date") else None,
            notes=request.form.get("notes"),
        )
        product_ids = request.form.getlist("item_product_id[]")
        quantities = request.form.getlist("item_quantity[]")
        prices = request.form.getlist("item_price[]")

        for i in range(len(product_ids)):
            if product_ids[i]:
                item = OrderItem(
                    product_id=int(product_ids[i]),
                    quantity=float(quantities[i]),
                    unit_price=float(prices[i]),
                )
                item.calculate()
                order.items.append(item)

        order.recalculate()
        db.session.add(order)
        db.session.commit()
        flash("Comanda creata!", "success")
        return redirect(url_for("orders.list_orders"))

    clients = Client.query.order_by(Client.name).all()
    products = Product.query.filter_by(is_active=True).order_by(Product.name).all()
    return render_template(
        "pages/order_form.html",
        order=None,
        clients=clients,
        products=products,
        order_number=generate_order_number(),
    )


@orders_bp.route("/<int:id>")
@login_required
def view_order(id):
    order = db.session.get(Order, id) or abort(404)
    return render_template("pages/order_view.html", order=order)


@orders_bp.route("/<int:id>/status/<status>", methods=["POST"])
@login_required
def update_status(id, status):
    order = db.session.get(Order, id) or abort(404)
    valid = ("new", "confirmed", "processing", "shipped", "delivered", "cancelled")
    if status in valid:
        order.status = status
        db.session.commit()
        flash(f"Status comanda: {status}", "success")
    return redirect(url_for("orders.view_order", id=id))


@orders_bp.route("/<int:id>/delete", methods=["POST"])
@login_required
def delete_order(id):
    order = db.session.get(Order, id) or abort(404)
    db.session.delete(order)
    db.session.commit()
    flash("Comanda stearsa!", "success")
    return redirect(url_for("orders.list_orders"))
