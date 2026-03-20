from flask import Blueprint, render_template
from modules import angajati, pontaj, concedii

main_bp = Blueprint('main', __name__)


@main_bp.route('/')
def dashboard():
    total_angajati = angajati.count_activi()
    prezenti_azi = pontaj.count_prezenti_azi()
    cereri_concediu = concedii.count_in_asteptare()
    angajati_recenti = angajati.get_all()[-5:]
    pontaj_recent = pontaj.get_all()[-10:]

    # Get employee names for pontaj
    ang_dict = {a['id']: f"{a['nume']} {a['prenume']}" for a in angajati.get_all()}
    for p in pontaj_recent:
        p['angajat_nume'] = ang_dict.get(p.get('angajat_id'), 'N/A')

    return render_template('dashboard.html',
                           total_angajati=total_angajati,
                           prezenti_azi=prezenti_azi,
                           cereri_concediu=cereri_concediu,
                           angajati_recenti=angajati_recenti,
                           pontaj_recent=pontaj_recent)
