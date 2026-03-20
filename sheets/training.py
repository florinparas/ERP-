"""
Foaia Training - Cursuri & formări profesionale
"""
from datetime import date

from config.hr_config import (
    TIPURI_TRAINING, STATUS_TRAINING,
    NUMBER_FORMAT_DATE, NUMBER_FORMAT_RON, NUMBER_FORMAT_INT, COLORS
)
from utils.helpers import (
    create_sheet_with_headers, add_dropdown_validation,
    add_status_conditional_formatting, add_date_expiry_formatting
)


def create_sheet(wb):
    """Creează foaia Training"""
    headers = [
        ("ID", 8),
        ("Denumire Curs", 25),
        ("Furnizor", 20),
        ("Tip", 14),
        ("Data Început", 14),
        ("Data Sfârșit", 14),
        ("Ore", 8),
        ("Cost (RON)", 14),
        ("ID Angajat", 12),
        ("Nume Angajat", 20),
        ("Status", 14),
        ("Certificare", 20),
        ("Data Expirare Certificare", 18),
    ]

    demo_data = [
        [
            "T001", "Management Avansat", "Schultz Consulting", "Extern",
            date(2025, 4, 10), date(2025, 4, 12), 24, 3500,
            "A001", None, "Planificat", "Da - Certificat Management", date(2027, 4, 12),
        ],
        [
            "T002", "Python pentru Analiză Date", "IT Academy", "Online",
            date(2025, 3, 1), date(2025, 3, 15), 40, 1200,
            "A002", None, "În Curs", "Da - Python Data Analysis", date(2027, 3, 15),
        ],
        [
            "T003", "Legislație Muncii 2025", "HR Club", "Conferință",
            date(2025, 5, 20), date(2025, 5, 20), 8, 500,
            "A003", None, "Planificat", "Nu", None,
        ],
        [
            "T004", "Excel Avansat & VBA", "Intern", "Intern",
            date(2025, 2, 15), date(2025, 2, 28), 16, 0,
            "A004", None, "Finalizat", "Da - Excel Advanced", date(2027, 2, 28),
        ],
    ]

    ws = create_sheet_with_headers(
        wb, "Training", "TRAINING & FORMĂRI PROFESIONALE", headers, demo_data
    )

    # VLOOKUP Nume Angajat
    for row in range(3, 3 + len(demo_data)):
        ws.cell(row=row, column=10).value = (
            f'=IFERROR(VLOOKUP(I{row},\'Angajați\'!$A:$C,2,0)&" "'
            f'&VLOOKUP(I{row},\'Angajați\'!$A:$C,3,0),"N/A")'
        )

    # Formatare date
    for col in [5, 6, 13]:
        for row in ws.iter_rows(min_row=3, max_row=ws.max_row,
                                 min_col=col, max_col=col):
            for cell in row:
                if cell.value:
                    cell.number_format = NUMBER_FORMAT_DATE

    # Formatare cost
    for row in ws.iter_rows(min_row=3, max_row=ws.max_row, min_col=8, max_col=8):
        for cell in row:
            cell.number_format = NUMBER_FORMAT_RON

    # Validări
    add_dropdown_validation(ws, "D3:D1000", TIPURI_TRAINING)
    add_dropdown_validation(ws, "K3:K1000", STATUS_TRAINING)

    # Formatare condițională - Status
    add_status_conditional_formatting(
        ws, "K", 3, 1000,
        {
            "Planificat": COLORS["light_blue"],
            "În Curs": COLORS["light_yellow"],
            "Finalizat": COLORS["light_green"],
            "Anulat": COLORS["light_red"],
        }
    )

    # Formatare condițională - expirare certificare
    add_date_expiry_formatting(ws, "M", 3, 1000, days_warning=90)

    return ws
