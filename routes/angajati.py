from flask import Blueprint, render_template, request, redirect, url_for, flash
from modules import angajati, departamente, pontaj, concedii, evaluari, documente
from config import TIP_CONTRACT_CHOICES, STATUS_ANGAJAT_CHOICES

angajati_bp = Blueprint('angajati', __name__)


@angajati_bp.route('/')
def lista():
    ang = angajati.get_all()
    deps = {d['id']: d['nume'] for d in departamente.get_all()}
    for a in ang:
        a['departament_nume'] = deps.get(a.get('departament_id'), '-')
    return render_template('angajati/lista.html', angajati=ang)


@angajati_bp.route('/nou', methods=['GET', 'POST'])
def nou():
    if request.method == 'POST':
        data = {
            'nume': request.form.get('nume', ''),
            'prenume': request.form.get('prenume', ''),
            'cnp': request.form.get('cnp', ''),
            'email': request.form.get('email', ''),
            'telefon': request.form.get('telefon', ''),
            'data_nasterii': request.form.get('data_nasterii', ''),
            'data_angajarii': request.form.get('data_angajarii', ''),
            'departament_id': int(request.form['departament_id']) if request.form.get('departament_id') else None,
            'functie': request.form.get('functie', ''),
            'tip_contract': request.form.get('tip_contract', ''),
            'status': request.form.get('status', 'activ'),
            'adresa': request.form.get('adresa', ''),
        }
        angajati.add(data)
        flash('Angajat adăugat cu succes!', 'success')
        return redirect(url_for('angajati.lista'))
    deps = departamente.get_all()
    return render_template('angajati/formular.html', angajat=None, departamente=deps,
                           tip_contract_choices=TIP_CONTRACT_CHOICES,
                           status_choices=STATUS_ANGAJAT_CHOICES)


@angajati_bp.route('/<int:angajat_id>')
def profil(angajat_id):
    ang = angajati.get_by_id(angajat_id)
    if not ang:
        flash('Angajatul nu a fost găsit.', 'error')
        return redirect(url_for('angajati.lista'))
    deps = {d['id']: d['nume'] for d in departamente.get_all()}
    ang['departament_nume'] = deps.get(ang.get('departament_id'), '-')
    ang_pontaj = pontaj.get_by_angajat(angajat_id)[-10:]
    ang_concedii = concedii.get_by_angajat(angajat_id)
    ang_evaluari = evaluari.get_by_angajat(angajat_id)
    ang_documente = documente.get_by_angajat(angajat_id)
    return render_template('angajati/profil.html', angajat=ang,
                           pontaj=ang_pontaj, concedii=ang_concedii,
                           evaluari=ang_evaluari, documente=ang_documente)


@angajati_bp.route('/<int:angajat_id>/editare', methods=['GET', 'POST'])
def editare(angajat_id):
    ang = angajati.get_by_id(angajat_id)
    if not ang:
        flash('Angajatul nu a fost găsit.', 'error')
        return redirect(url_for('angajati.lista'))
    if request.method == 'POST':
        data = {
            'nume': request.form.get('nume', ''),
            'prenume': request.form.get('prenume', ''),
            'cnp': request.form.get('cnp', ''),
            'email': request.form.get('email', ''),
            'telefon': request.form.get('telefon', ''),
            'data_nasterii': request.form.get('data_nasterii', ''),
            'data_angajarii': request.form.get('data_angajarii', ''),
            'departament_id': int(request.form['departament_id']) if request.form.get('departament_id') else None,
            'functie': request.form.get('functie', ''),
            'tip_contract': request.form.get('tip_contract', ''),
            'status': request.form.get('status', 'activ'),
            'adresa': request.form.get('adresa', ''),
        }
        angajati.modify(angajat_id, data)
        flash('Angajat actualizat cu succes!', 'success')
        return redirect(url_for('angajati.profil', angajat_id=angajat_id))
    deps = departamente.get_all()
    return render_template('angajati/formular.html', angajat=ang, departamente=deps,
                           tip_contract_choices=TIP_CONTRACT_CHOICES,
                           status_choices=STATUS_ANGAJAT_CHOICES)


@angajati_bp.route('/<int:angajat_id>/stergere', methods=['POST'])
def stergere(angajat_id):
    angajati.remove(angajat_id)
    flash('Angajat șters cu succes!', 'success')
    return redirect(url_for('angajati.lista'))
