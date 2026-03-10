from datetime import datetime, timezone
from erp_app import db


class Order(db.Model):
    __tablename__ = "orders"

    id = db.Column(db.Integer, primary_key=True)
    number = db.Column(db.String(50), unique=True, nullable=False)
    client_id = db.Column(db.Integer, db.ForeignKey("clients.id"), nullable=False)
    date = db.Column(db.Date, default=lambda: datetime.now(timezone.utc).date())
    delivery_date = db.Column(db.Date)
    status = db.Column(db.String(20), default="new")  # new, confirmed, processing, shipped, delivered, cancelled
    total = db.Column(db.Float, default=0.0)
    notes = db.Column(db.Text)
    created_at = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))

    items = db.relationship("OrderItem", backref="order", cascade="all, delete-orphan")

    def recalculate(self):
        self.total = sum(item.total for item in self.items)

    def to_dict(self):
        return {
            "id": self.id,
            "number": self.number,
            "client_id": self.client_id,
            "client_name": self.client.name if self.client else "",
            "date": self.date.isoformat() if self.date else None,
            "delivery_date": self.delivery_date.isoformat() if self.delivery_date else None,
            "status": self.status,
            "total": self.total,
            "items": [item.to_dict() for item in self.items],
        }


class OrderItem(db.Model):
    __tablename__ = "order_items"

    id = db.Column(db.Integer, primary_key=True)
    order_id = db.Column(db.Integer, db.ForeignKey("orders.id"), nullable=False)
    product_id = db.Column(db.Integer, db.ForeignKey("products.id"), nullable=False)
    quantity = db.Column(db.Float, nullable=False, default=1.0)
    unit_price = db.Column(db.Float, nullable=False, default=0.0)
    total = db.Column(db.Float, default=0.0)

    product = db.relationship("Product")

    def calculate(self):
        self.total = self.quantity * self.unit_price

    def to_dict(self):
        return {
            "id": self.id,
            "product_id": self.product_id,
            "product_name": self.product.name if self.product else "",
            "quantity": self.quantity,
            "unit_price": self.unit_price,
            "total": self.total,
        }
