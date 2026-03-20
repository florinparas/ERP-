from flask import Blueprint, render_template, request, redirect, url_for, flash
from modules import departamente, angajati

departamente_bp = Blueprint('departamente', __name__)


@departamente_bp.route('/')
def lista():
    deps = departamente.get_all()
    ang_dict = {a['id']: f"{a['nume']} {a['prenume']}" for a in angajati.get_all()}
    for d in deps:
        d['manager_nume'] = ang_dict.get(d.get('manager_id'), '-')
    return render_template('departamente/lista.html', departamente=deps)


@departamente_bp.route('/nou', methods=['GET', 'POST'])
def nou():
    if request.method == 'POST':
        data = {
            'nume': request.form.get('nume', ''),
            'cod': request.form.get('cod', ''),
            'manager_id': int(request.form['manager_id']) if request.form.get('manager_id') else None,
            'parinte_id': int(request.form['parinte_id']) if request.form.get('parinte_id') else None,
            'descriere': request.form.get('descriere', ''),
        }
        departamente.add(data)
        flash('Departament adăugat cu succes!', 'success')
        return redirect(url_for('departamente.lista'))
    ang = angajati.get_all()
    deps = departamente.get_all()
    return render_template('departamente/formular.html', departament=None, angajati=ang, departamente=deps)


@departamente_bp.route('/<int:dept_id>/editare', methods=['GET', 'POST'])
def editare(dept_id):
    dept = departamente.get_by_id(dept_id)
    if not dept:
        flash('Departamentul nu a fost găsit.', 'error')
        return redirect(url_for('departamente.lista'))
    if request.method == 'POST':
        data = {
            'nume': request.form.get('nume', ''),
            'cod': request.form.get('cod', ''),
            'manager_id': int(request.form['manager_id']) if request.form.get('manager_id') else None,
            'parinte_id': int(request.form['parinte_id']) if request.form.get('parinte_id') else None,
            'descriere': request.form.get('descriere', ''),
        }
        departamente.modify(dept_id, data)
        flash('Departament actualizat cu succes!', 'success')
        return redirect(url_for('departamente.lista'))
    ang = angajati.get_all()
    deps = [d for d in departamente.get_all() if d['id'] != dept_id]
    return render_template('departamente/formular.html', departament=dept, angajati=ang, departamente=deps)


@departamente_bp.route('/<int:dept_id>/stergere', methods=['POST'])
def stergere(dept_id):
    departamente.remove(dept_id)
    flash('Departament șters cu succes!', 'success')
    return redirect(url_for('departamente.lista'))
