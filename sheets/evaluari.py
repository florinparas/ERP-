"""
Foaia Evaluări - Evaluări performanță
"""
from datetime import date

from config.hr_config import (
    NUMBER_FORMAT_DATE, COLORS
)
from utils.helpers import (
    create_sheet_with_headers, add_dropdown_validation,
    add_number_validation, add_score_conditional_formatting,
    add_date_expiry_formatting
)


def create_sheet(wb):
    """Creează foaia Evaluări"""
    headers = [
        ("ID", 8),
        ("ID Angajat", 12),
        ("Nume Angajat", 20),
        ("Perioada Evaluată", 18),
        ("Data Evaluării", 14),
        ("Evaluator", 18),
        ("Obiective", 30),
        ("Realizări", 30),
        ("Scor Obiective (1-5)", 16),
        ("Competențe Tehnice (1-5)", 18),
        ("Competențe Soft (1-5)", 18),
        ("Scor Final", 12),
        ("Plan Dezvoltare", 25),
        ("Următoarea Evaluare", 16),
    ]

    demo_data = [
        [
            "E001", "A002", None, "S1 2025", date(2025, 6, 30),
            "Popescu Ion", "Livrare proiect X, Mentorat 2 juniori",
            "Proiect livrat la termen, 1 junior mentorat", 4, 5, 4,
            None, "Certificare cloud, leadership training",
            date(2025, 12, 31),
        ],
        [
            "E002", "A003", None, "S1 2025", date(2025, 6, 30),
            "Popescu Ion", "Implementare sistem evaluare, Reducere turnover 5%",
            "Sistem implementat parțial", 3, 3, 4,
            None, "Curs management HR avansat",
            date(2025, 12, 31),
        ],
        [
            "E003", "A005", None, "S1 2025", date(2025, 6, 30),
            "Popescu Ion", "Creștere vânzări 15%, Noi clienți",
            "Creștere 20% vânzări, 5 clienți noi", 5, 4, 4,
            None, "Curs negociere avansată",
            date(2025, 12, 31),
        ],
    ]

    ws = create_sheet_with_headers(
        wb, "Evaluări", "EVALUĂRI PERFORMANȚĂ", headers, demo_data
    )

    # VLOOKUP Nume
    for row in range(3, 3 + len(demo_data)):
        ws.cell(row=row, column=3).value = (
            f'=IFERROR(VLOOKUP(B{row},\'Angajați\'!$A:$C,2,0)&" "'
            f'&VLOOKUP(B{row},\'Angajați\'!$A:$C,3,0),"N/A")'
        )

    # Formula Scor Final (media celor 3 scoruri)
    for row in range(3, 3 + len(demo_data)):
        ws.cell(row=row, column=12).value = (
            f'=IF(AND(I{row}<>"",J{row}<>"",K{row}<>""),'
            f'ROUND((I{row}+J{row}+K{row})/3,1),"")'
        )

    # Formatare date
    for row in ws.iter_rows(min_row=3, max_row=ws.max_row, min_col=5, max_col=5):
        for cell in row:
            if cell.value:
                cell.number_format = NUMBER_FORMAT_DATE
    for row in ws.iter_rows(min_row=3, max_row=ws.max_row, min_col=14, max_col=14):
        for cell in row:
            if cell.value:
                cell.number_format = NUMBER_FORMAT_DATE

    # Validări scor 1-5
    add_number_validation(ws, "I3:I1000", 1, 5)
    add_number_validation(ws, "J3:J1000", 1, 5)
    add_number_validation(ws, "K3:K1000", 1, 5)

    # Formatare condițională scoruri
    add_score_conditional_formatting(ws, "I", 3, 1000)
    add_score_conditional_formatting(ws, "J", 3, 1000)
    add_score_conditional_formatting(ws, "K", 3, 1000)
    add_score_conditional_formatting(ws, "L", 3, 1000)

    # Formatare condițională - evaluări scadente
    add_date_expiry_formatting(ws, "N", 3, 1000, days_warning=30)

    return ws
