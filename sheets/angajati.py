"""
Foaia Angajați - Evidența personalului
"""
from datetime import datetime

from config.hr_config import (
    ANGAJATI_DEMO, STATUS_ANGAJAT, TIPURI_CONTRACT, SEXE,
    NUMBER_FORMAT_DATE, COLORS
)
from utils.helpers import (
    create_sheet_with_headers, add_dropdown_validation,
    add_status_conditional_formatting
)
from utils.styles import get_border


def create_sheet(wb):
    """Creează foaia Angajați"""
    headers = [
        ("ID", 8),
        ("Nume", 15),
        ("Prenume", 15),
        ("CNP", 16),
        ("Data Nașterii", 14),
        ("Sex", 6),
        ("Adresă", 30),
        ("Telefon", 14),
        ("Email", 25),
        ("Data Angajării", 14),
        ("Departament", 20),
        ("Funcție", 22),
        ("Manager Direct", 18),
        ("Tip Contract", 18),
        ("Status", 12),
        ("Data Ieșire", 14),
        ("Motiv Ieșire", 20),
        ("Observații", 25),
    ]

    # Pregătire date demo
    data = []
    for ang in ANGAJATI_DEMO:
        row = [
            ang["id"],
            ang["nume"],
            ang["prenume"],
            ang["cnp"],
            datetime.strptime(ang["data_nasterii"], "%Y-%m-%d").date(),
            ang["sex"],
            ang["adresa"],
            ang["telefon"],
            ang["email"],
            datetime.strptime(ang["data_angajarii"], "%Y-%m-%d").date(),
            ang["departament"],
            ang["functie"],
            ang["manager"],
            ang["tip_contract"],
            ang["status"],
            None,  # Data ieșire
            None,  # Motiv ieșire
            None,  # Observații
        ]
        data.append(row)

    ws = create_sheet_with_headers(
        wb, "Angajați", "EVIDENȚA ANGAJAȚILOR", headers, data
    )

    # Formatare date
    for row in ws.iter_rows(min_row=3, max_row=ws.max_row, min_col=5, max_col=5):
        for cell in row:
            if cell.value:
                cell.number_format = NUMBER_FORMAT_DATE
    for row in ws.iter_rows(min_row=3, max_row=ws.max_row, min_col=10, max_col=10):
        for cell in row:
            if cell.value:
                cell.number_format = NUMBER_FORMAT_DATE
    for row in ws.iter_rows(min_row=3, max_row=ws.max_row, min_col=16, max_col=16):
        for cell in row:
            if cell.value:
                cell.number_format = NUMBER_FORMAT_DATE

    # Validări dropdown
    add_dropdown_validation(ws, "F3:F1000", SEXE)
    add_dropdown_validation(ws, "K3:K1000", "=Departamente!$B$3:$B$50")
    add_dropdown_validation(ws, "L3:L1000", "=Departamente!$I$3:$I$50")
    add_dropdown_validation(ws, "N3:N1000", TIPURI_CONTRACT)
    add_dropdown_validation(ws, "O3:O1000", STATUS_ANGAJAT)

    # Formatare condițională pentru Status
    add_status_conditional_formatting(
        ws, "O", 3, 1000,
        {
            "Activ": COLORS["light_green"],
            "Inactiv": COLORS["light_red"],
            "Suspendat": COLORS["light_yellow"],
        }
    )

    return ws
