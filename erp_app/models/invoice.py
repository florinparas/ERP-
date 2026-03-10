from datetime import datetime, timezone
from erp_app import db


class Invoice(db.Model):
    __tablename__ = "invoices"

    id = db.Column(db.Integer, primary_key=True)
    number = db.Column(db.String(50), unique=True, nullable=False)
    client_id = db.Column(db.Integer, db.ForeignKey("clients.id"), nullable=False)
    date = db.Column(db.Date, default=lambda: datetime.now(timezone.utc).date())
    due_date = db.Column(db.Date)
    status = db.Column(db.String(20), default="draft")  # draft, sent, paid, overdue, cancelled
    subtotal = db.Column(db.Float, default=0.0)
    vat_total = db.Column(db.Float, default=0.0)
    total = db.Column(db.Float, default=0.0)
    notes = db.Column(db.Text)
    created_at = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))

    items = db.relationship("InvoiceItem", backref="invoice", cascade="all, delete-orphan")

    def recalculate(self):
        self.subtotal = sum(item.total for item in self.items)
        self.vat_total = sum(item.vat_amount for item in self.items)
        self.total = self.subtotal + self.vat_total

    def to_dict(self):
        return {
            "id": self.id,
            "number": self.number,
            "client_id": self.client_id,
            "client_name": self.client.name if self.client else "",
            "date": self.date.isoformat() if self.date else None,
            "due_date": self.due_date.isoformat() if self.due_date else None,
            "status": self.status,
            "subtotal": self.subtotal,
            "vat_total": self.vat_total,
            "total": self.total,
            "items": [item.to_dict() for item in self.items],
        }


class InvoiceItem(db.Model):
    __tablename__ = "invoice_items"

    id = db.Column(db.Integer, primary_key=True)
    invoice_id = db.Column(db.Integer, db.ForeignKey("invoices.id"), nullable=False)
    product_id = db.Column(db.Integer, db.ForeignKey("products.id"))
    description = db.Column(db.String(300), nullable=False)
    quantity = db.Column(db.Float, nullable=False, default=1.0)
    unit_price = db.Column(db.Float, nullable=False, default=0.0)
    vat_rate = db.Column(db.Float, default=19.0)
    total = db.Column(db.Float, default=0.0)
    vat_amount = db.Column(db.Float, default=0.0)

    product = db.relationship("Product")

    def calculate(self):
        self.total = self.quantity * self.unit_price
        self.vat_amount = self.total * self.vat_rate / 100

    def to_dict(self):
        return {
            "id": self.id,
            "description": self.description,
            "quantity": self.quantity,
            "unit_price": self.unit_price,
            "vat_rate": self.vat_rate,
            "total": self.total,
            "vat_amount": self.vat_amount,
        }
