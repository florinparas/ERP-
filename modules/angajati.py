from modules.excel_manager import read_all, read_by_id, create, update, delete, filter_by
from config import ANGAJATI_FILE


def get_all():
    return read_all(ANGAJATI_FILE)


def get_by_id(angajat_id):
    return read_by_id(ANGAJATI_FILE, angajat_id)


def get_activi():
    return filter_by(ANGAJATI_FILE, status='activ')


def get_by_departament(dept_id):
    return filter_by(ANGAJATI_FILE, departament_id=dept_id)


def add(data):
    return create(ANGAJATI_FILE, data)


def modify(angajat_id, data):
    update(ANGAJATI_FILE, angajat_id, data)


def remove(angajat_id):
    delete(ANGAJATI_FILE, angajat_id)


def count_activi():
    return len(get_activi())


def get_nume_complet(angajat):
    if angajat:
        return f"{angajat.get('nume', '')} {angajat.get('prenume', '')}"
    return ''
