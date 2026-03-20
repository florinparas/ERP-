"""
Configurare modul HR - Constante, texte, culori, liste dropdown
ERP Management Personal - România
"""

# ============================================================================
# PARAMETRI FISCALI ROMÂNIA 2025
# ============================================================================
CAS_RATE = 0.25          # Contribuție asigurări sociale (pensie)
CASS_RATE = 0.10         # Contribuție asigurări sociale de sănătate
TAX_RATE = 0.10          # Impozit pe venit (flat tax)
CAM_RATE = 0.0225        # Contribuția asiguratorie pentru muncă (angajator)

SALARIU_MINIM_BRUT = 4050  # RON, de la 1 ianuarie 2025
SALARIU_MINIM_CONSTRUCTII = 4582  # RON, sector construcții

ZILE_CONCEDIU_STANDARD = 21  # Zile lucrătoare/an
ORE_LUCRU_ZI = 8             # Normă completă

# Deducere personală de bază (simplificată)
DEDUCERE_PERSONALA_BAZA = 300  # RON

# ============================================================================
# SĂRBĂTORI LEGALE ROMÂNIA 2025-2026
# ============================================================================
SARBATORI_LEGALE = {
    2025: [
        "01-01",  # Anul Nou
        "02-01",  # Anul Nou (ziua 2)
        "24-01",  # Ziua Unirii Principatelor Române
        "18-04",  # Vinerea Mare (Ortodoxă)
        "20-04",  # Paștele Ortodox
        "21-04",  # Paștele Ortodox (ziua 2)
        "01-05",  # Ziua Muncii
        "01-06",  # Ziua Copilului
        "08-06",  # Rusaliile
        "09-06",  # Rusaliile (ziua 2)
        "15-08",  # Adormirea Maicii Domnului
        "30-11",  # Sfântul Andrei
        "01-12",  # Ziua Națională a României
        "25-12",  # Crăciunul
        "26-12",  # Crăciunul (ziua 2)
    ],
    2026: [
        "01-01",  # Anul Nou
        "02-01",  # Anul Nou (ziua 2)
        "24-01",  # Ziua Unirii Principatelor Române
        "10-04",  # Vinerea Mare (Ortodoxă)
        "12-04",  # Paștele Ortodox
        "13-04",  # Paștele Ortodox (ziua 2)
        "01-05",  # Ziua Muncii
        "01-06",  # Ziua Copilului
        "31-05",  # Rusaliile
        "01-06",  # Rusaliile (ziua 2)
        "15-08",  # Adormirea Maicii Domnului
        "30-11",  # Sfântul Andrei
        "01-12",  # Ziua Națională a României
        "25-12",  # Crăciunul
        "26-12",  # Crăciunul (ziua 2)
    ],
}

# ============================================================================
# CULORI TEMĂ (hex fără #, pentru openpyxl)
# ============================================================================
COLORS = {
    "header_bg": "1F4E79",       # Albastru închis
    "header_font": "FFFFFF",     # Alb
    "row_alt": "D6E4F0",         # Albastru deschis (zebra)
    "row_normal": "FFFFFF",      # Alb
    "accent": "2E75B6",          # Albastru accent
    "success": "70AD47",         # Verde
    "warning": "FFC000",         # Galben/Portocaliu
    "danger": "FF0000",          # Roșu
    "info": "5B9BD5",            # Albastru info
    "light_green": "E2EFDA",     # Verde deschis
    "light_red": "FCE4EC",       # Roșu deschis
    "light_yellow": "FFF9C4",    # Galben deschis
    "light_blue": "DEEAF6",      # Albastru deschis
    "light_orange": "FFE0B2",    # Portocaliu deschis
    "border": "B4C6E7",          # Gri-albastru bordură
    "title_bg": "0D3B66",        # Albastru foarte închis (titluri)
}

# ============================================================================
# LISTE DROPDOWN
# ============================================================================
STATUS_ANGAJAT = ["Activ", "Inactiv", "Suspendat"]

TIPURI_CONTRACT = [
    "CIM Nedeterminat",
    "CIM Determinat",
    "Part-time",
    "Convenție Civilă",
    "Internship",
]

TIPURI_CONCEDIU = [
    "CO - Concediu Odihnă",
    "CM - Concediu Medical",
    "CFS - Concediu Fără Salariu",
    "CCC - Concediu Creștere Copil",
    "CE - Concediu Eveniment",
    "CP - Concediu Paternitate",
    "CM - Concediu Maternitate",
]

STATUS_CONCEDIU = ["Aprobat", "Respins", "În Așteptare"]

CODURI_PONTAJ = [
    "P",    # Prezent
    "CO",   # Concediu odihnă
    "CM",   # Concediu medical
    "A",    # Absent nemotivat
    "AM",   # Absent motivat
    "LS",   # Zi liberă / sărbătoare
    "OS",   # Ore suplimentare
    "TP",   # Telemuncă / remote
    "D",    # Delegație
]

SEXE = ["M", "F"]

NIVEL_FUNCTIE = ["Junior", "Mid", "Senior", "Lead", "Manager", "Director"]

STATUS_RECRUTARE = [
    "Nou",
    "Screening CV",
    "Interviu Telefonic",
    "Interviu Tehnic",
    "Interviu Final",
    "Ofertă",
    "Angajat",
    "Respins",
    "Retras",
]

SURSE_RECRUTARE = [
    "LinkedIn",
    "Site Companie",
    "Referral",
    "Agenție",
    "eJobs",
    "BestJobs",
    "Hipo",
    "Altele",
]

TIPURI_TRAINING = ["Intern", "Extern", "Online", "Conferință", "Workshop"]

STATUS_TRAINING = ["Planificat", "În Curs", "Finalizat", "Anulat"]

SCOR_EVALUARE = [1, 2, 3, 4, 5]

# ============================================================================
# DATE DEMO - DEPARTAMENTE
# ============================================================================
DEPARTAMENTE_DEMO = [
    {"id": "D001", "denumire": "Management", "manager": "Popescu Ion", "locatie": "București", "buget": 50000},
    {"id": "D002", "denumire": "IT & Dezvoltare", "manager": "Ionescu Maria", "locatie": "București", "buget": 120000},
    {"id": "D003", "denumire": "Resurse Umane", "manager": "Georgescu Ana", "locatie": "București", "buget": 40000},
    {"id": "D004", "denumire": "Financiar-Contabil", "manager": "Dumitrescu Pavel", "locatie": "București", "buget": 60000},
    {"id": "D005", "denumire": "Vânzări", "manager": "Marinescu Elena", "locatie": "Cluj-Napoca", "buget": 80000},
    {"id": "D006", "denumire": "Marketing", "manager": "Constantinescu Dan", "locatie": "București", "buget": 55000},
    {"id": "D007", "denumire": "Logistică", "manager": "Vasilescu Mihai", "locatie": "Timișoara", "buget": 70000},
]

FUNCTII_DEMO = [
    {"id": "F001", "denumire": "Director General", "dept": "Management", "nivel": "Director", "sal_min": 15000, "sal_max": 30000},
    {"id": "F002", "denumire": "Developer Software", "dept": "IT & Dezvoltare", "nivel": "Mid", "sal_min": 7000, "sal_max": 14000},
    {"id": "F003", "denumire": "Specialist HR", "dept": "Resurse Umane", "nivel": "Mid", "sal_min": 5000, "sal_max": 9000},
    {"id": "F004", "denumire": "Contabil", "dept": "Financiar-Contabil", "nivel": "Mid", "sal_min": 5500, "sal_max": 10000},
    {"id": "F005", "denumire": "Manager Vânzări", "dept": "Vânzări", "nivel": "Manager", "sal_min": 8000, "sal_max": 15000},
    {"id": "F006", "denumire": "Specialist Marketing", "dept": "Marketing", "nivel": "Mid", "sal_min": 5000, "sal_max": 9000},
    {"id": "F007", "denumire": "Operator Logistică", "dept": "Logistică", "nivel": "Junior", "sal_min": 4050, "sal_max": 6000},
    {"id": "F008", "denumire": "Team Lead Dezvoltare", "dept": "IT & Dezvoltare", "nivel": "Lead", "sal_min": 12000, "sal_max": 20000},
    {"id": "F009", "denumire": "Analist Business", "dept": "IT & Dezvoltare", "nivel": "Mid", "sal_min": 6000, "sal_max": 11000},
    {"id": "F010", "denumire": "Asistent Manager", "dept": "Management", "nivel": "Junior", "sal_min": 4500, "sal_max": 7000},
]

# ============================================================================
# DATE DEMO - ANGAJAȚI
# ============================================================================
ANGAJATI_DEMO = [
    {
        "id": "A001", "nume": "Popescu", "prenume": "Ion", "cnp": "1800515400123",
        "data_nasterii": "1980-05-15", "sex": "M", "adresa": "Str. Victoriei 10, București",
        "telefon": "0721000001", "email": "ion.popescu@company.ro",
        "data_angajarii": "2015-03-01", "departament": "Management",
        "functie": "Director General", "manager": "-", "tip_contract": "CIM Nedeterminat",
        "status": "Activ",
    },
    {
        "id": "A002", "nume": "Ionescu", "prenume": "Maria", "cnp": "2850720400234",
        "data_nasterii": "1985-07-20", "sex": "F", "adresa": "Bd. Unirii 25, București",
        "telefon": "0721000002", "email": "maria.ionescu@company.ro",
        "data_angajarii": "2017-06-15", "departament": "IT & Dezvoltare",
        "functie": "Team Lead Dezvoltare", "manager": "Popescu Ion", "tip_contract": "CIM Nedeterminat",
        "status": "Activ",
    },
    {
        "id": "A003", "nume": "Georgescu", "prenume": "Ana", "cnp": "2900312400345",
        "data_nasterii": "1990-03-12", "sex": "F", "adresa": "Str. Florilor 5, București",
        "telefon": "0721000003", "email": "ana.georgescu@company.ro",
        "data_angajarii": "2019-01-10", "departament": "Resurse Umane",
        "functie": "Specialist HR", "manager": "Popescu Ion", "tip_contract": "CIM Nedeterminat",
        "status": "Activ",
    },
    {
        "id": "A004", "nume": "Dumitrescu", "prenume": "Pavel", "cnp": "1880225400456",
        "data_nasterii": "1988-02-25", "sex": "M", "adresa": "Str. Libertății 18, București",
        "telefon": "0721000004", "email": "pavel.dumitrescu@company.ro",
        "data_angajarii": "2016-09-01", "departament": "Financiar-Contabil",
        "functie": "Contabil", "manager": "Popescu Ion", "tip_contract": "CIM Nedeterminat",
        "status": "Activ",
    },
    {
        "id": "A005", "nume": "Marinescu", "prenume": "Elena", "cnp": "2920618400567",
        "data_nasterii": "1992-06-18", "sex": "F", "adresa": "Str. Primăverii 30, Cluj-Napoca",
        "telefon": "0721000005", "email": "elena.marinescu@company.ro",
        "data_angajarii": "2020-02-01", "departament": "Vânzări",
        "functie": "Manager Vânzări", "manager": "Popescu Ion", "tip_contract": "CIM Nedeterminat",
        "status": "Activ",
    },
]

# ============================================================================
# CONFIGURARE FOI EXCEL
# ============================================================================
SHEET_ORDER = [
    "Dashboard",
    "Angajați",
    "Contracte",
    "Departamente",
    "Pontaj",
    "Concedii",
    "Salarizare",
    "Evaluări",
    "Training",
    "Recrutare",
    "Configurare",
]

# Format monetar românesc
NUMBER_FORMAT_RON = '#,##0.00 "RON"'
NUMBER_FORMAT_PERCENT = '0.00%'
NUMBER_FORMAT_DATE = 'DD.MM.YYYY'
NUMBER_FORMAT_INT = '#,##0'
