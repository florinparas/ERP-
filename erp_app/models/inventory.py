from datetime import datetime, timezone
from erp_app import db


class InventoryMovement(db.Model):
    __tablename__ = "inventory_movements"

    id = db.Column(db.Integer, primary_key=True)
    product_id = db.Column(db.Integer, db.ForeignKey("products.id"), nullable=False)
    movement_type = db.Column(db.String(20), nullable=False)  # in, out, adjustment
    quantity = db.Column(db.Float, nullable=False)
    reference = db.Column(db.String(100))  # order number, invoice number, etc.
    notes = db.Column(db.Text)
    created_at = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))

    product = db.relationship("Product")

    def to_dict(self):
        return {
            "id": self.id,
            "product_id": self.product_id,
            "product_name": self.product.name if self.product else "",
            "movement_type": self.movement_type,
            "quantity": self.quantity,
            "reference": self.reference,
            "notes": self.notes,
            "created_at": self.created_at.isoformat() if self.created_at else None,
        }
