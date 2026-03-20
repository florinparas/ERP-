from datetime import date
from modules.excel_manager import read_all, read_by_id, create, update, delete, filter_by
from config import EVALUARI_FILE


def get_all():
    return read_all(EVALUARI_FILE)


def get_by_id(evaluare_id):
    return read_by_id(EVALUARI_FILE, evaluare_id)


def get_by_angajat(angajat_id):
    return filter_by(EVALUARI_FILE, angajat_id=int(angajat_id))


def add(data):
    data['data_evaluare'] = date.today().isoformat()
    if 'status' not in data:
        data['status'] = 'draft'
    try:
        data['scor_general'] = float(data.get('scor_general', 0))
    except (ValueError, TypeError):
        data['scor_general'] = 0
    return create(EVALUARI_FILE, data)


def modify(evaluare_id, data):
    try:
        data['scor_general'] = float(data.get('scor_general', 0))
    except (ValueError, TypeError):
        data['scor_general'] = 0
    update(EVALUARI_FILE, evaluare_id, data)


def remove(evaluare_id):
    delete(EVALUARI_FILE, evaluare_id)
