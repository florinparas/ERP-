from flask import Blueprint, render_template, request, redirect, url_for, flash
from modules import concedii, angajati
from config import TIP_CONCEDIU_CHOICES

concedii_bp = Blueprint('concedii', __name__)


@concedii_bp.route('/')
def lista():
    records = concedii.get_all()
    ang_dict = {a['id']: f"{a['nume']} {a['prenume']}" for a in angajati.get_all()}
    for r in records:
        r['angajat_nume'] = ang_dict.get(r.get('angajat_id'), 'N/A')
        r['aprobat_de_nume'] = ang_dict.get(r.get('aprobat_de'), '-')
    records.reverse()
    return render_template('concedii/lista.html', concedii=records)


@concedii_bp.route('/cerere', methods=['GET', 'POST'])
def cerere():
    if request.method == 'POST':
        data = {
            'angajat_id': int(request.form.get('angajat_id', 0)),
            'tip_concediu': request.form.get('tip_concediu', ''),
            'data_inceput': request.form.get('data_inceput', ''),
            'data_sfarsit': request.form.get('data_sfarsit', ''),
            'motiv': request.form.get('motiv', ''),
        }
        concedii.add_cerere(data)
        flash('Cerere de concediu trimisă cu succes!', 'success')
        return redirect(url_for('concedii.lista'))
    ang = angajati.get_activi()
    return render_template('concedii/cerere.html', angajati=ang,
                           tip_choices=TIP_CONCEDIU_CHOICES)


@concedii_bp.route('/<int:concediu_id>/aprobare', methods=['POST'])
def aprobare(concediu_id):
    aprobat_de = request.form.get('aprobat_de')
    concedii.aproba(concediu_id, int(aprobat_de) if aprobat_de else None)
    flash('Concediu aprobat!', 'success')
    return redirect(url_for('concedii.lista'))


@concedii_bp.route('/<int:concediu_id>/respingere', methods=['POST'])
def respingere(concediu_id):
    aprobat_de = request.form.get('aprobat_de')
    concedii.respinge(concediu_id, int(aprobat_de) if aprobat_de else None)
    flash('Concediu respins.', 'info')
    return redirect(url_for('concedii.lista'))
