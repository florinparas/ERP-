"""
Foaia Concedii - Management concedii & zile libere
"""
from datetime import date

from config.hr_config import (
    TIPURI_CONCEDIU, STATUS_CONCEDIU, NUMBER_FORMAT_DATE,
    NUMBER_FORMAT_INT, ZILE_CONCEDIU_STANDARD, COLORS
)
from utils.helpers import (
    create_sheet_with_headers, add_dropdown_validation,
    add_status_conditional_formatting
)


def create_sheet(wb):
    """Creează foaia Concedii"""
    headers = [
        ("ID", 8),
        ("ID Angajat", 12),
        ("Nume Angajat", 20),
        ("Tip Concediu", 25),
        ("Data Început", 14),
        ("Data Sfârșit", 14),
        ("Zile", 8),
        ("Status", 14),
        ("Aprobat De", 18),
        ("Zile Disponibile/An", 16),
        ("Zile Folosite", 14),
        ("Zile Rămase", 12),
    ]

    demo_data = [
        [
            "L001", "A002", None,
            "CO - Concediu Odihnă", date(2025, 7, 14), date(2025, 7, 25),
            None, "Aprobat", "Popescu Ion",
            ZILE_CONCEDIU_STANDARD, None, None,
        ],
        [
            "L002", "A003", None,
            "CO - Concediu Odihnă", date(2025, 8, 4), date(2025, 8, 15),
            None, "În Așteptare", "",
            ZILE_CONCEDIU_STANDARD, None, None,
        ],
        [
            "L003", "A004", None,
            "CM - Concediu Medical", date(2025, 3, 10), date(2025, 3, 14),
            None, "Aprobat", "Popescu Ion",
            ZILE_CONCEDIU_STANDARD, None, None,
        ],
    ]

    ws = create_sheet_with_headers(
        wb, "Concedii", "MANAGEMENT CONCEDII", headers, demo_data
    )

    # Formule VLOOKUP pentru Nume Angajat
    for row in range(3, 3 + len(demo_data)):
        ws.cell(row=row, column=3).value = (
            f'=IFERROR(VLOOKUP(B{row},\'Angajați\'!$A:$C,2,0)&" "'
            f'&VLOOKUP(B{row},\'Angajați\'!$A:$C,3,0),"N/A")'
        )

    # Formula Zile (diferența între date, doar zile lucrătoare simplificate)
    for row in range(3, 3 + len(demo_data)):
        ws.cell(row=row, column=7).value = (
            f'=IF(AND(E{row}<>"",F{row}<>""),NETWORKDAYS(E{row},F{row}),"")'
        )

    # Formula Zile Folosite (SUMIFS pe angajat, an curent, status aprobat)
    for row in range(3, 3 + len(demo_data)):
        ws.cell(row=row, column=11).value = (
            f'=SUMPRODUCT((B$3:B$1000=B{row})'
            f'*(YEAR(E$3:E$1000)=YEAR(TODAY()))'
            f'*(H$3:H$1000="Aprobat")'
            f'*(G$3:G$1000))'
        )

    # Formula Zile Rămase
    for row in range(3, 3 + len(demo_data)):
        ws.cell(row=row, column=12).value = f'=IF(J{row}<>"",J{row}-K{row},"")'

    # Formatare date
    for row in ws.iter_rows(min_row=3, max_row=ws.max_row, min_col=5, max_col=6):
        for cell in row:
            if cell.value:
                cell.number_format = NUMBER_FORMAT_DATE

    # Validări
    add_dropdown_validation(ws, "D3:D1000", TIPURI_CONCEDIU)
    add_dropdown_validation(ws, "H3:H1000", STATUS_CONCEDIU)

    # Formatare condițională - Status concediu
    add_status_conditional_formatting(
        ws, "H", 3, 1000,
        {
            "Aprobat": COLORS["light_green"],
            "Respins": COLORS["light_red"],
            "În Așteptare": COLORS["light_yellow"],
        }
    )

    return ws
