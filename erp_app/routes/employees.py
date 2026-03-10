from datetime import datetime
from flask import Blueprint, render_template, request, redirect, url_for, flash, abort
from flask_login import login_required
from erp_app import db
from erp_app.models.employee import Employee

employees_bp = Blueprint("employees", __name__)


@employees_bp.route("/")
@login_required
def list_employees():
    page = request.args.get("page", 1, type=int)
    search = request.args.get("search", "")
    department = request.args.get("department", "")
    query = Employee.query
    if search:
        query = query.filter(
            db.or_(
                Employee.first_name.ilike(f"%{search}%"),
                Employee.last_name.ilike(f"%{search}%"),
            )
        )
    if department:
        query = query.filter_by(department=department)
    employees = query.order_by(Employee.last_name).paginate(page=page, per_page=20)
    departments = db.session.query(Employee.department).distinct().all()
    return render_template(
        "pages/employees.html",
        employees=employees,
        search=search,
        departments=[d[0] for d in departments if d[0]],
        selected_department=department,
    )


@employees_bp.route("/add", methods=["GET", "POST"])
@login_required
def add_employee():
    if request.method == "POST":
        employee = Employee(
            first_name=request.form["first_name"],
            last_name=request.form["last_name"],
            email=request.form.get("email"),
            phone=request.form.get("phone"),
            position=request.form.get("position"),
            department=request.form.get("department"),
            salary=float(request.form.get("salary", 0)),
            hire_date=datetime.strptime(request.form["hire_date"], "%Y-%m-%d").date() if request.form.get("hire_date") else None,
            cnp=request.form.get("cnp"),
            address=request.form.get("address"),
        )
        db.session.add(employee)
        db.session.commit()
        flash("Angajat adaugat!", "success")
        return redirect(url_for("employees.list_employees"))
    return render_template("pages/employee_form.html", employee=None)


@employees_bp.route("/<int:id>/edit", methods=["GET", "POST"])
@login_required
def edit_employee(id):
    employee = db.session.get(Employee, id) or abort(404)
    if request.method == "POST":
        employee.first_name = request.form["first_name"]
        employee.last_name = request.form["last_name"]
        employee.email = request.form.get("email")
        employee.phone = request.form.get("phone")
        employee.position = request.form.get("position")
        employee.department = request.form.get("department")
        employee.salary = float(request.form.get("salary", 0))
        employee.hire_date = datetime.strptime(request.form["hire_date"], "%Y-%m-%d").date() if request.form.get("hire_date") else None
        employee.cnp = request.form.get("cnp")
        employee.address = request.form.get("address")
        db.session.commit()
        flash("Angajat actualizat!", "success")
        return redirect(url_for("employees.list_employees"))
    return render_template("pages/employee_form.html", employee=employee)


@employees_bp.route("/<int:id>/delete", methods=["POST"])
@login_required
def delete_employee(id):
    employee = db.session.get(Employee, id) or abort(404)
    db.session.delete(employee)
    db.session.commit()
    flash("Angajat sters!", "success")
    return redirect(url_for("employees.list_employees"))
