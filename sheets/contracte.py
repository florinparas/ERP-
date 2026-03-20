"""
Foaia Contracte - Contracte de muncă
"""
from datetime import datetime, date

from config.hr_config import (
    TIPURI_CONTRACT, STATUS_ANGAJAT, NUMBER_FORMAT_DATE,
    NUMBER_FORMAT_RON, NUMBER_FORMAT_INT, COLORS
)
from utils.helpers import (
    create_sheet_with_headers, add_dropdown_validation,
    add_status_conditional_formatting, add_date_expiry_formatting,
    vlookup_formula
)


def create_sheet(wb):
    """Creează foaia Contracte"""
    headers = [
        ("ID Contract", 12),
        ("ID Angajat", 12),
        ("Nume Angajat", 20),
        ("Tip Contract", 20),
        ("Număr Contract", 16),
        ("Data Început", 14),
        ("Data Sfârșit", 14),
        ("Normă (ore/zi)", 14),
        ("Salariu Brut (RON)", 18),
        ("Clauze Speciale", 25),
        ("Status", 12),
        ("Acte Adiționale", 20),
    ]

    # Date demo
    demo_data = [
        [
            "C001", "A001", None, "CIM Nedeterminat", "CIM-001/2015",
            date(2015, 3, 1), None, 8, 25000, "Clauză confidențialitate",
            "Activ", "",
        ],
        [
            "C002", "A002", None, "CIM Nedeterminat", "CIM-002/2017",
            date(2017, 6, 15), None, 8, 14000, "",
            "Activ", "AA1 - mărire salariu 01.2024",
        ],
        [
            "C003", "A003", None, "CIM Nedeterminat", "CIM-003/2019",
            date(2019, 1, 10), None, 8, 7500, "",
            "Activ", "",
        ],
        [
            "C004", "A004", None, "CIM Nedeterminat", "CIM-004/2016",
            date(2016, 9, 1), None, 8, 9000, "Clauză neconcurență",
            "Activ", "",
        ],
        [
            "C005", "A005", None, "CIM Nedeterminat", "CIM-005/2020",
            date(2020, 2, 1), None, 8, 12000, "",
            "Activ", "",
        ],
    ]

    ws = create_sheet_with_headers(
        wb, "Contracte", "CONTRACTE DE MUNCĂ", headers, demo_data
    )

    # Formule VLOOKUP pentru Nume Angajat (coloana C)
    for row in range(3, 3 + len(demo_data)):
        ws.cell(row=row, column=3).value = (
            f'=IFERROR(VLOOKUP(B{row},\'Angajați\'!$A:$C,2,0)&" "'
            f'&VLOOKUP(B{row},\'Angajați\'!$A:$C,3,0),"N/A")'
        )

    # Formatare date
    for row in ws.iter_rows(min_row=3, max_row=ws.max_row, min_col=6, max_col=7):
        for cell in row:
            if cell.value:
                cell.number_format = NUMBER_FORMAT_DATE

    # Formatare salariu
    for row in ws.iter_rows(min_row=3, max_row=ws.max_row, min_col=9, max_col=9):
        for cell in row:
            cell.number_format = NUMBER_FORMAT_RON

    # Validări
    add_dropdown_validation(ws, "D3:D1000", TIPURI_CONTRACT)
    add_dropdown_validation(ws, "K3:K1000", STATUS_ANGAJAT)

    # Formatare condițională - Status
    add_status_conditional_formatting(
        ws, "K", 3, 1000,
        {
            "Activ": COLORS["light_green"],
            "Inactiv": COLORS["light_red"],
            "Suspendat": COLORS["light_yellow"],
        }
    )

    # Formatare condițională - Data Sfârșit (expirare contract)
    add_date_expiry_formatting(ws, "G", 3, 1000, days_warning=30)

    return ws
