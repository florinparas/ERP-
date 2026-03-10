from flask import Flask
from flask_sqlalchemy import SQLAlchemy
from flask_migrate import Migrate
from flask_login import LoginManager
from config import Config

db = SQLAlchemy()
migrate = Migrate()
login_manager = LoginManager()
login_manager.login_view = "auth.login"


def create_app(config_class=Config):
    app = Flask(__name__)
    app.config.from_object(config_class)

    db.init_app(app)
    migrate.init_app(app, db)
    login_manager.init_app(app)

    from erp_app.routes.auth import auth_bp
    from erp_app.routes.dashboard import dashboard_bp
    from erp_app.routes.clients import clients_bp
    from erp_app.routes.products import products_bp
    from erp_app.routes.invoices import invoices_bp
    from erp_app.routes.orders import orders_bp
    from erp_app.routes.employees import employees_bp
    from erp_app.routes.inventory import inventory_bp
    from erp_app.routes.office import office_bp

    app.register_blueprint(auth_bp)
    app.register_blueprint(dashboard_bp)
    app.register_blueprint(clients_bp, url_prefix="/clients")
    app.register_blueprint(products_bp, url_prefix="/products")
    app.register_blueprint(invoices_bp, url_prefix="/invoices")
    app.register_blueprint(orders_bp, url_prefix="/orders")
    app.register_blueprint(employees_bp, url_prefix="/employees")
    app.register_blueprint(inventory_bp, url_prefix="/inventory")
    app.register_blueprint(office_bp, url_prefix="/office")

    with app.app_context():
        db.create_all()

    return app
