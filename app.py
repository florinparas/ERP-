from erp_app import create_app, db
from erp_app.models.user import User

app = create_app()


@app.cli.command("init-db")
def init_db():
    """Initializeaza baza de date si creeaza utilizatorul admin."""
    db.create_all()
    if not User.query.filter_by(username="admin").first():
        admin = User(username="admin", email="admin@erp.local", role="admin")
        admin.set_password("admin123")
        db.session.add(admin)
        db.session.commit()
        print("Baza de date initializata. Admin: admin / admin123")
    else:
        print("Baza de date deja initializata.")


if __name__ == "__main__":
    with app.app_context():
        db.create_all()
        if not User.query.filter_by(username="admin").first():
            admin = User(username="admin", email="admin@erp.local", role="admin")
            admin.set_password("admin123")
            db.session.add(admin)
            db.session.commit()
    app.run(debug=True, host="0.0.0.0", port=5000)
