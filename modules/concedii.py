from datetime import datetime, date
from modules.excel_manager import read_all, read_by_id, create, update, filter_by
from config import CONCEDII_FILE


def get_all():
    return read_all(CONCEDII_FILE)


def get_by_id(concediu_id):
    return read_by_id(CONCEDII_FILE, concediu_id)


def get_by_angajat(angajat_id):
    return filter_by(CONCEDII_FILE, angajat_id=int(angajat_id))


def get_in_asteptare():
    return filter_by(CONCEDII_FILE, status='in_asteptare')


def count_in_asteptare():
    return len(get_in_asteptare())


def add_cerere(data):
    try:
        d1 = datetime.strptime(str(data['data_inceput']), '%Y-%m-%d').date()
        d2 = datetime.strptime(str(data['data_sfarsit']), '%Y-%m-%d').date()
        data['zile_total'] = (d2 - d1).days + 1
    except (ValueError, TypeError, KeyError):
        data['zile_total'] = 0
    data['status'] = 'in_asteptare'
    data['data_cerere'] = date.today().isoformat()
    return create(CONCEDII_FILE, data)


def aproba(concediu_id, aprobat_de):
    update(CONCEDII_FILE, concediu_id, {
        'status': 'aprobat',
        'aprobat_de': aprobat_de,
    })


def respinge(concediu_id, aprobat_de):
    update(CONCEDII_FILE, concediu_id, {
        'status': 'respins',
        'aprobat_de': aprobat_de,
    })
