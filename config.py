import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, 'data')
UPLOAD_DIR = os.path.join(BASE_DIR, 'uploads')

SECRET_KEY = 'erp-secret-key-change-in-production'

# Excel file paths
ANGAJATI_FILE = os.path.join(DATA_DIR, 'angajati.xlsx')
DEPARTAMENTE_FILE = os.path.join(DATA_DIR, 'departamente.xlsx')
PONTAJ_FILE = os.path.join(DATA_DIR, 'pontaj.xlsx')
CONCEDII_FILE = os.path.join(DATA_DIR, 'concedii.xlsx')
EVALUARI_FILE = os.path.join(DATA_DIR, 'evaluari.xlsx')
DOCUMENTE_FILE = os.path.join(DATA_DIR, 'documente.xlsx')

# Excel schemas (column headers for each file)
SCHEMAS = {
    'angajati': ['id', 'nume', 'prenume', 'cnp', 'email', 'telefon',
                 'data_nasterii', 'data_angajarii', 'departament_id',
                 'functie', 'tip_contract', 'status', 'adresa'],
    'departamente': ['id', 'nume', 'cod', 'manager_id', 'parinte_id', 'descriere'],
    'pontaj': ['id', 'angajat_id', 'data', 'ora_intrare', 'ora_iesire',
               'ore_lucrate', 'tip', 'observatii'],
    'concedii': ['id', 'angajat_id', 'tip_concediu', 'data_inceput',
                 'data_sfarsit', 'zile_total', 'status', 'aprobat_de',
                 'data_cerere', 'motiv'],
    'evaluari': ['id', 'angajat_id', 'evaluator_id', 'data_evaluare',
                 'perioada', 'scor_general', 'obiective', 'puncte_forte',
                 'puncte_slabe', 'comentarii', 'status'],
    'documente': ['id', 'angajat_id', 'nume_fisier', 'cale_fisier',
                  'tip_document', 'data_incarcare', 'descriere'],
}

# Constants
TIP_CONTRACT_CHOICES = ['CIM', 'PFA', 'Colaborare']
STATUS_ANGAJAT_CHOICES = ['activ', 'inactiv', 'suspendat']
TIP_CONCEDIU_CHOICES = ['odihna', 'medical', 'fara_plata', 'maternitate', 'studii']
STATUS_CONCEDIU_CHOICES = ['in_asteptare', 'aprobat', 'respins']
TIP_PONTAJ_CHOICES = ['normal', 'ore_suplimentare', 'tura_noapte']
TIP_DOCUMENT_CHOICES = ['contract', 'act_identitate', 'certificat', 'altele']
STATUS_EVALUARE_CHOICES = ['draft', 'finalizat']
