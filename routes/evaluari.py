from flask import Blueprint, render_template, request, redirect, url_for, flash
from modules import evaluari, angajati
from config import STATUS_EVALUARE_CHOICES

evaluari_bp = Blueprint('evaluari', __name__)


@evaluari_bp.route('/')
def lista():
    records = evaluari.get_all()
    ang_dict = {a['id']: f"{a['nume']} {a['prenume']}" for a in angajati.get_all()}
    for r in records:
        r['angajat_nume'] = ang_dict.get(r.get('angajat_id'), 'N/A')
        r['evaluator_nume'] = ang_dict.get(r.get('evaluator_id'), 'N/A')
    records.reverse()
    return render_template('evaluari/lista.html', evaluari=records)


@evaluari_bp.route('/nou', methods=['GET', 'POST'])
def nou():
    if request.method == 'POST':
        data = {
            'angajat_id': int(request.form.get('angajat_id', 0)),
            'evaluator_id': int(request.form.get('evaluator_id', 0)),
            'perioada': request.form.get('perioada', ''),
            'scor_general': request.form.get('scor_general', 0),
            'obiective': request.form.get('obiective', ''),
            'puncte_forte': request.form.get('puncte_forte', ''),
            'puncte_slabe': request.form.get('puncte_slabe', ''),
            'comentarii': request.form.get('comentarii', ''),
            'status': request.form.get('status', 'draft'),
        }
        evaluari.add(data)
        flash('Evaluare adăugată cu succes!', 'success')
        return redirect(url_for('evaluari.lista'))
    ang = angajati.get_all()
    return render_template('evaluari/formular.html', evaluare=None, angajati=ang,
                           status_choices=STATUS_EVALUARE_CHOICES)


@evaluari_bp.route('/<int:evaluare_id>')
def vizualizare(evaluare_id):
    ev = evaluari.get_by_id(evaluare_id)
    if not ev:
        flash('Evaluarea nu a fost găsită.', 'error')
        return redirect(url_for('evaluari.lista'))
    ang_dict = {a['id']: f"{a['nume']} {a['prenume']}" for a in angajati.get_all()}
    ev['angajat_nume'] = ang_dict.get(ev.get('angajat_id'), 'N/A')
    ev['evaluator_nume'] = ang_dict.get(ev.get('evaluator_id'), 'N/A')
    return render_template('evaluari/raport.html', evaluare=ev)


@evaluari_bp.route('/<int:evaluare_id>/editare', methods=['GET', 'POST'])
def editare(evaluare_id):
    ev = evaluari.get_by_id(evaluare_id)
    if not ev:
        flash('Evaluarea nu a fost găsită.', 'error')
        return redirect(url_for('evaluari.lista'))
    if request.method == 'POST':
        data = {
            'angajat_id': int(request.form.get('angajat_id', 0)),
            'evaluator_id': int(request.form.get('evaluator_id', 0)),
            'perioada': request.form.get('perioada', ''),
            'scor_general': request.form.get('scor_general', 0),
            'obiective': request.form.get('obiective', ''),
            'puncte_forte': request.form.get('puncte_forte', ''),
            'puncte_slabe': request.form.get('puncte_slabe', ''),
            'comentarii': request.form.get('comentarii', ''),
            'status': request.form.get('status', 'draft'),
        }
        evaluari.modify(evaluare_id, data)
        flash('Evaluare actualizată cu succes!', 'success')
        return redirect(url_for('evaluari.vizualizare', evaluare_id=evaluare_id))
    ang = angajati.get_all()
    return render_template('evaluari/formular.html', evaluare=ev, angajati=ang,
                           status_choices=STATUS_EVALUARE_CHOICES)


@evaluari_bp.route('/<int:evaluare_id>/stergere', methods=['POST'])
def stergere(evaluare_id):
    evaluari.remove(evaluare_id)
    flash('Evaluare ștearsă cu succes!', 'success')
    return redirect(url_for('evaluari.lista'))
