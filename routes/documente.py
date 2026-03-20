import os
from flask import Blueprint, render_template, request, redirect, url_for, flash, send_from_directory
from modules import documente, angajati
from config import TIP_DOCUMENT_CHOICES, UPLOAD_DIR

documente_bp = Blueprint('documente', __name__)


@documente_bp.route('/')
def lista():
    records = documente.get_all()
    ang_dict = {a['id']: f"{a['nume']} {a['prenume']}" for a in angajati.get_all()}
    for r in records:
        r['angajat_nume'] = ang_dict.get(r.get('angajat_id'), 'N/A')
    records.reverse()
    return render_template('documente/lista.html', documente=records)


@documente_bp.route('/upload', methods=['GET', 'POST'])
def upload():
    if request.method == 'POST':
        file = request.files.get('fisier')
        angajat_id = request.form.get('angajat_id')
        tip_document = request.form.get('tip_document', 'altele')
        descriere = request.form.get('descriere', '')

        if not file or file.filename == '':
            flash('Selectați un fișier.', 'error')
            return redirect(url_for('documente.upload'))

        result = documente.upload(file, angajat_id, tip_document, descriere)
        if result:
            flash('Document încărcat cu succes!', 'success')
        else:
            flash('Tip de fișier neacceptat. Extensii permise: pdf, doc, docx, jpg, jpeg, png, xls, xlsx', 'error')
        return redirect(url_for('documente.lista'))

    ang = angajati.get_all()
    return render_template('documente/upload.html', angajati=ang,
                           tip_choices=TIP_DOCUMENT_CHOICES)


@documente_bp.route('/<int:doc_id>/descarca')
def descarca(doc_id):
    doc = documente.get_by_id(doc_id)
    if not doc or not doc.get('cale_fisier'):
        flash('Documentul nu a fost găsit.', 'error')
        return redirect(url_for('documente.lista'))
    return send_from_directory(UPLOAD_DIR, doc['cale_fisier'],
                               as_attachment=True,
                               download_name=doc.get('nume_fisier', doc['cale_fisier']))


@documente_bp.route('/<int:doc_id>/stergere', methods=['POST'])
def stergere(doc_id):
    documente.remove(doc_id)
    flash('Document șters cu succes!', 'success')
    return redirect(url_for('documente.lista'))
