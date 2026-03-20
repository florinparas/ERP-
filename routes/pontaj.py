from flask import Blueprint, render_template, request, redirect, url_for, flash
from modules import pontaj, angajati
from config import TIP_PONTAJ_CHOICES
from datetime import datetime

pontaj_bp = Blueprint('pontaj', __name__)


@pontaj_bp.route('/')
def lista():
    records = pontaj.get_all()
    ang_dict = {a['id']: f"{a['nume']} {a['prenume']}" for a in angajati.get_all()}
    for r in records:
        r['angajat_nume'] = ang_dict.get(r.get('angajat_id'), 'N/A')
    records.reverse()
    return render_template('pontaj/lista.html', pontaj=records)


@pontaj_bp.route('/inregistrare', methods=['GET', 'POST'])
def inregistrare():
    if request.method == 'POST':
        action = request.form.get('action')
        if action == 'check_in':
            angajat_id = request.form.get('angajat_id')
            tip = request.form.get('tip', 'normal')
            observatii = request.form.get('observatii', '')
            pontaj.check_in(angajat_id, tip, observatii)
            flash('Check-in înregistrat cu succes!', 'success')
        elif action == 'check_out':
            pontaj_id = request.form.get('pontaj_id')
            pontaj.check_out(pontaj_id)
            flash('Check-out înregistrat cu succes!', 'success')
        return redirect(url_for('pontaj.inregistrare'))

    ang = angajati.get_activi()
    prezenti = pontaj.get_prezenti_azi()
    ang_dict = {a['id']: f"{a['nume']} {a['prenume']}" for a in angajati.get_all()}
    for p in prezenti:
        p['angajat_nume'] = ang_dict.get(p.get('angajat_id'), 'N/A')
    return render_template('pontaj/inregistrare.html', angajati=ang, prezenti=prezenti,
                           tip_choices=TIP_PONTAJ_CHOICES)


@pontaj_bp.route('/raport')
def raport():
    luna = request.args.get('luna', datetime.now().month)
    an = request.args.get('an', datetime.now().year)
    angajat_id = request.args.get('angajat_id', '')
    records = []
    if angajat_id:
        records = pontaj.get_pontaj_luna(angajat_id, luna, an)
    ang = angajati.get_all()
    ang_dict = {a['id']: f"{a['nume']} {a['prenume']}" for a in ang}
    for r in records:
        r['angajat_nume'] = ang_dict.get(r.get('angajat_id'), 'N/A')
    total_ore = sum(float(r.get('ore_lucrate') or 0) for r in records)
    return render_template('pontaj/raport.html', pontaj=records, angajati=ang,
                           luna=int(luna), an=int(an), angajat_id=angajat_id,
                           total_ore=round(total_ore, 2))
