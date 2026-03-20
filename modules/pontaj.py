from datetime import datetime, date
from modules.excel_manager import read_all, read_by_id, create, update, filter_by
from config import PONTAJ_FILE


def get_all():
    return read_all(PONTAJ_FILE)


def get_by_id(pontaj_id):
    return read_by_id(PONTAJ_FILE, pontaj_id)


def get_prezenti_azi():
    today = date.today().isoformat()
    entries = filter_by(PONTAJ_FILE, data=today)
    return [e for e in entries if e.get('ora_iesire') is None]


def get_pontaj_luna(angajat_id, luna, an):
    records = read_all(PONTAJ_FILE)
    results = []
    for r in records:
        if str(r.get('angajat_id', '')) == str(angajat_id):
            data_val = r.get('data', '')
            if data_val:
                try:
                    d = datetime.strptime(str(data_val), '%Y-%m-%d').date()
                    if d.month == int(luna) and d.year == int(an):
                        results.append(r)
                except (ValueError, TypeError):
                    pass
    return results


def check_in(angajat_id, tip='normal', observatii=''):
    now = datetime.now()
    data = {
        'angajat_id': int(angajat_id),
        'data': now.strftime('%Y-%m-%d'),
        'ora_intrare': now.strftime('%H:%M'),
        'ora_iesire': None,
        'ore_lucrate': None,
        'tip': tip,
        'observatii': observatii,
    }
    return create(PONTAJ_FILE, data)


def check_out(pontaj_id):
    record = get_by_id(pontaj_id)
    if not record:
        return False
    now = datetime.now()
    ora_intrare = record.get('ora_intrare', '')
    try:
        intrare = datetime.strptime(f"{record['data']} {ora_intrare}", '%Y-%m-%d %H:%M')
        ore_lucrate = round((now - intrare).total_seconds() / 3600, 2)
    except (ValueError, TypeError):
        ore_lucrate = 0
    update(PONTAJ_FILE, pontaj_id, {
        'ora_iesire': now.strftime('%H:%M'),
        'ore_lucrate': ore_lucrate,
    })
    return True


def count_prezenti_azi():
    today = date.today().isoformat()
    entries = filter_by(PONTAJ_FILE, data=today)
    return len(entries)


def get_by_angajat(angajat_id):
    return filter_by(PONTAJ_FILE, angajat_id=int(angajat_id))
