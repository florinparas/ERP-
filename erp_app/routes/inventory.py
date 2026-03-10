from flask import Blueprint, render_template, request, redirect, url_for, flash, abort
from flask_login import login_required
from erp_app import db
from erp_app.models.inventory import InventoryMovement
from erp_app.models.product import Product

inventory_bp = Blueprint("inventory", __name__)


@inventory_bp.route("/")
@login_required
def list_movements():
    page = request.args.get("page", 1, type=int)
    product_id = request.args.get("product_id", type=int)
    query = InventoryMovement.query
    if product_id:
        query = query.filter_by(product_id=product_id)
    movements = query.order_by(InventoryMovement.created_at.desc()).paginate(page=page, per_page=20)
    products = Product.query.order_by(Product.name).all()
    low_stock = Product.query.filter(Product.stock_quantity <= Product.min_stock).all()
    return render_template(
        "pages/inventory.html",
        movements=movements,
        products=products,
        low_stock=low_stock,
        selected_product=product_id,
    )


@inventory_bp.route("/add", methods=["GET", "POST"])
@login_required
def add_movement():
    if request.method == "POST":
        product_id = int(request.form["product_id"])
        movement_type = request.form["movement_type"]
        quantity = float(request.form["quantity"])

        movement = InventoryMovement(
            product_id=product_id,
            movement_type=movement_type,
            quantity=quantity,
            reference=request.form.get("reference"),
            notes=request.form.get("notes"),
        )

        product = db.session.get(Product, product_id)
        if product:
            if movement_type == "in":
                product.stock_quantity += quantity
            elif movement_type == "out":
                product.stock_quantity -= quantity
            elif movement_type == "adjustment":
                product.stock_quantity = quantity

        db.session.add(movement)
        db.session.commit()
        flash("Miscare stoc inregistrata!", "success")
        return redirect(url_for("inventory.list_movements"))

    products = Product.query.filter_by(is_active=True).order_by(Product.name).all()
    return render_template("pages/inventory_form.html", products=products)
