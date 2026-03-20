from modules.excel_manager import read_all, read_by_id, create, update, delete
from config import DEPARTAMENTE_FILE


def get_all():
    return read_all(DEPARTAMENTE_FILE)


def get_by_id(dept_id):
    return read_by_id(DEPARTAMENTE_FILE, dept_id)


def add(data):
    return create(DEPARTAMENTE_FILE, data)


def modify(dept_id, data):
    update(DEPARTAMENTE_FILE, dept_id, data)


def remove(dept_id):
    delete(DEPARTAMENTE_FILE, dept_id)
