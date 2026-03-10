from datetime import datetime, timezone
from erp_app import db


class Product(db.Model):
    __tablename__ = "products"

    id = db.Column(db.Integer, primary_key=True)
    code = db.Column(db.String(50), unique=True, nullable=False)
    name = db.Column(db.String(200), nullable=False)
    description = db.Column(db.Text)
    category = db.Column(db.String(100))
    unit = db.Column(db.String(20), default="buc")  # buc, kg, l, m
    price = db.Column(db.Float, nullable=False, default=0.0)
    cost = db.Column(db.Float, default=0.0)
    vat_rate = db.Column(db.Float, default=19.0)  # TVA standard Romania
    stock_quantity = db.Column(db.Float, default=0.0)
    min_stock = db.Column(db.Float, default=0.0)
    is_active = db.Column(db.Boolean, default=True)
    created_at = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))

    def to_dict(self):
        return {
            "id": self.id,
            "code": self.code,
            "name": self.name,
            "description": self.description,
            "category": self.category,
            "unit": self.unit,
            "price": self.price,
            "cost": self.cost,
            "vat_rate": self.vat_rate,
            "stock_quantity": self.stock_quantity,
            "min_stock": self.min_stock,
            "is_active": self.is_active,
        }
