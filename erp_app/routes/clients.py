from flask import Blueprint, render_template, request, redirect, url_for, flash
from flask_login import login_required
from erp_app import db
from erp_app.models.client import Client

clients_bp = Blueprint("clients", __name__)


@clients_bp.route("/")
@login_required
def list_clients():
    page = request.args.get("page", 1, type=int)
    search = request.args.get("search", "")
    query = Client.query
    if search:
        query = query.filter(
            db.or_(
                Client.name.ilike(f"%{search}%"),
                Client.company.ilike(f"%{search}%"),
                Client.email.ilike(f"%{search}%"),
            )
        )
    clients = query.order_by(Client.name).paginate(page=page, per_page=20)
    return render_template("pages/clients.html", clients=clients, search=search)


@clients_bp.route("/add", methods=["GET", "POST"])
@login_required
def add_client():
    if request.method == "POST":
        client = Client(
            name=request.form["name"],
            email=request.form.get("email"),
            phone=request.form.get("phone"),
            company=request.form.get("company"),
            cui=request.form.get("cui"),
            address=request.form.get("address"),
            city=request.form.get("city"),
            country=request.form.get("country", "Romania"),
            notes=request.form.get("notes"),
        )
        db.session.add(client)
        db.session.commit()
        flash("Client adaugat cu succes!", "success")
        return redirect(url_for("clients.list_clients"))
    return render_template("pages/client_form.html", client=None)


@clients_bp.route("/<int:id>/edit", methods=["GET", "POST"])
@login_required
def edit_client(id):
    client = db.session.get(Client, id) or abort(404)
    if request.method == "POST":
        client.name = request.form["name"]
        client.email = request.form.get("email")
        client.phone = request.form.get("phone")
        client.company = request.form.get("company")
        client.cui = request.form.get("cui")
        client.address = request.form.get("address")
        client.city = request.form.get("city")
        client.country = request.form.get("country", "Romania")
        client.notes = request.form.get("notes")
        db.session.commit()
        flash("Client actualizat!", "success")
        return redirect(url_for("clients.list_clients"))
    return render_template("pages/client_form.html", client=client)


@clients_bp.route("/<int:id>/delete", methods=["POST"])
@login_required
def delete_client(id):
    client = db.session.get(Client, id) or abort(404)
    db.session.delete(client)
    db.session.commit()
    flash("Client sters!", "success")
    return redirect(url_for("clients.list_clients"))
