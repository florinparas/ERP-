from flask import Blueprint, render_template
from flask_login import login_required
from erp_app import db
from erp_app.models import Client, Product, Invoice, Order, Employee

dashboard_bp = Blueprint("dashboard", __name__)


@dashboard_bp.route("/")
@login_required
def index():
    stats = {
        "total_clients": Client.query.count(),
        "total_products": Product.query.count(),
        "total_invoices": Invoice.query.count(),
        "total_orders": Order.query.count(),
        "total_employees": Employee.query.count(),
        "unpaid_invoices": Invoice.query.filter_by(status="sent").count(),
        "pending_orders": Order.query.filter(
            Order.status.in_(["new", "confirmed", "processing"])
        ).count(),
        "low_stock_products": Product.query.filter(
            Product.stock_quantity <= Product.min_stock
        ).count(),
        "revenue": db.session.query(
            db.func.coalesce(db.func.sum(Invoice.total), 0)
        ).filter_by(status="paid").scalar(),
    }
    recent_invoices = Invoice.query.order_by(Invoice.created_at.desc()).limit(5).all()
    recent_orders = Order.query.order_by(Order.created_at.desc()).limit(5).all()
    return render_template(
        "pages/dashboard.html",
        stats=stats,
        recent_invoices=recent_invoices,
        recent_orders=recent_orders,
    )
