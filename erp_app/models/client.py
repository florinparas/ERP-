from datetime import datetime, timezone
from erp_app import db


class Client(db.Model):
    __tablename__ = "clients"

    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(200), nullable=False)
    email = db.Column(db.String(120))
    phone = db.Column(db.String(20))
    company = db.Column(db.String(200))
    cui = db.Column(db.String(20))  # Cod Unic de Inregistrare
    address = db.Column(db.Text)
    city = db.Column(db.String(100))
    country = db.Column(db.String(100), default="Romania")
    notes = db.Column(db.Text)
    created_at = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))

    invoices = db.relationship("Invoice", backref="client", lazy="dynamic")
    orders = db.relationship("Order", backref="client", lazy="dynamic")

    def to_dict(self):
        return {
            "id": self.id,
            "name": self.name,
            "email": self.email,
            "phone": self.phone,
            "company": self.company,
            "cui": self.cui,
            "address": self.address,
            "city": self.city,
            "country": self.country,
            "notes": self.notes,
            "created_at": self.created_at.isoformat() if self.created_at else None,
        }
