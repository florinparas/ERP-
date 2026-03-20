import os
from datetime import date
from werkzeug.utils import secure_filename
from modules.excel_manager import read_all, read_by_id, create, delete, filter_by
from config import DOCUMENTE_FILE, UPLOAD_DIR


ALLOWED_EXTENSIONS = {'pdf', 'doc', 'docx', 'jpg', 'jpeg', 'png', 'xls', 'xlsx'}


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def get_all():
    return read_all(DOCUMENTE_FILE)


def get_by_id(doc_id):
    return read_by_id(DOCUMENTE_FILE, doc_id)


def get_by_angajat(angajat_id):
    return filter_by(DOCUMENTE_FILE, angajat_id=int(angajat_id))


def upload(file, angajat_id, tip_document, descriere=''):
    if not file or not allowed_file(file.filename):
        return None
    filename = secure_filename(file.filename)
    # Add timestamp to avoid collisions
    name, ext = os.path.splitext(filename)
    unique_name = f"{name}_{date.today().isoformat()}_{angajat_id}{ext}"
    filepath = os.path.join(UPLOAD_DIR, unique_name)
    os.makedirs(UPLOAD_DIR, exist_ok=True)
    file.save(filepath)

    data = {
        'angajat_id': int(angajat_id),
        'nume_fisier': file.filename,
        'cale_fisier': unique_name,
        'tip_document': tip_document,
        'data_incarcare': date.today().isoformat(),
        'descriere': descriere,
    }
    return create(DOCUMENTE_FILE, data)


def remove(doc_id):
    doc = get_by_id(doc_id)
    if doc and doc.get('cale_fisier'):
        filepath = os.path.join(UPLOAD_DIR, doc['cale_fisier'])
        if os.path.exists(filepath):
            os.remove(filepath)
    delete(DOCUMENTE_FILE, doc_id)
