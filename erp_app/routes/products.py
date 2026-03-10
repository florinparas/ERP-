from flask import Blueprint, render_template, request, redirect, url_for, flash, abort
from flask_login import login_required
from erp_app import db
from erp_app.models.product import Product

products_bp = Blueprint("products", __name__)


@products_bp.route("/")
@login_required
def list_products():
    page = request.args.get("page", 1, type=int)
    search = request.args.get("search", "")
    category = request.args.get("category", "")
    query = Product.query
    if search:
        query = query.filter(
            db.or_(
                Product.name.ilike(f"%{search}%"),
                Product.code.ilike(f"%{search}%"),
            )
        )
    if category:
        query = query.filter_by(category=category)
    products = query.order_by(Product.name).paginate(page=page, per_page=20)
    categories = db.session.query(Product.category).distinct().all()
    return render_template(
        "pages/products.html",
        products=products,
        search=search,
        categories=[c[0] for c in categories if c[0]],
        selected_category=category,
    )


@products_bp.route("/add", methods=["GET", "POST"])
@login_required
def add_product():
    if request.method == "POST":
        product = Product(
            code=request.form["code"],
            name=request.form["name"],
            description=request.form.get("description"),
            category=request.form.get("category"),
            unit=request.form.get("unit", "buc"),
            price=float(request.form.get("price", 0)),
            cost=float(request.form.get("cost", 0)),
            vat_rate=float(request.form.get("vat_rate", 19)),
            stock_quantity=float(request.form.get("stock_quantity", 0)),
            min_stock=float(request.form.get("min_stock", 0)),
        )
        db.session.add(product)
        db.session.commit()
        flash("Produs adaugat cu succes!", "success")
        return redirect(url_for("products.list_products"))
    return render_template("pages/product_form.html", product=None)


@products_bp.route("/<int:id>/edit", methods=["GET", "POST"])
@login_required
def edit_product(id):
    product = db.session.get(Product, id) or abort(404)
    if request.method == "POST":
        product.code = request.form["code"]
        product.name = request.form["name"]
        product.description = request.form.get("description")
        product.category = request.form.get("category")
        product.unit = request.form.get("unit", "buc")
        product.price = float(request.form.get("price", 0))
        product.cost = float(request.form.get("cost", 0))
        product.vat_rate = float(request.form.get("vat_rate", 19))
        product.min_stock = float(request.form.get("min_stock", 0))
        db.session.commit()
        flash("Produs actualizat!", "success")
        return redirect(url_for("products.list_products"))
    return render_template("pages/product_form.html", product=product)


@products_bp.route("/<int:id>/delete", methods=["POST"])
@login_required
def delete_product(id):
    product = db.session.get(Product, id) or abort(404)
    db.session.delete(product)
    db.session.commit()
    flash("Produs sters!", "success")
    return redirect(url_for("products.list_products"))
