"""
Foaia Documente Angajați - Management documente personal
"""
from datetime import date

from config.hr_config import NUMBER_FORMAT_DATE, COLORS
from utils.helpers import (
    create_sheet_with_headers, add_dropdown_validation,
    add_status_conditional_formatting, add_date_expiry_formatting
)


TIPURI_DOCUMENT = [
    "Carte Identitate",
    "Pașaport",
    "Permis Conducere",
    "Diplomă Studii",
    "Certificat Profesional",
    "Certificat Medical",
    "Cazier Judiciar",
    "Contract de Muncă",
    "Act Adițional",
    "Fișa Postului",
    "Adeverință Vechime",
    "Declarație GDPR",
    "SSM - Fișă Instruire",
    "Altele",
]

STATUS_DOCUMENT = ["Valid", "Expirat", "Expiră Curând", "Lipsă"]


def create_sheet(wb):
    """Creează foaia Documente Angajați"""
    headers = [
        ("ID", 8),
        ("ID Angajat", 12),
        ("Nume Angajat", 20),
        ("Tip Document", 22),
        ("Nr. Document", 16),
        ("Data Emiterii", 14),
        ("Data Expirării", 14),
        ("Emitent", 20),
        ("Status", 14),
        ("Observații", 25),
    ]

    demo_data = [
        [
            "DOC001", "A001", None, "Carte Identitate", "RX123456",
            date(2020, 5, 10), date(2030, 5, 10), "SPCLEP Sector 1",
            None, "",
        ],
        [
            "DOC002", "A001", None, "Fișa Postului", "FP-001/2015",
            date(2015, 3, 1), None, "SC Example SRL",
            None, "Actualizată 2023",
        ],
        [
            "DOC003", "A002", None, "Carte Identitate", "RX234567",
            date(2019, 8, 15), date(2029, 8, 15), "SPCLEP Sector 2",
            None, "",
        ],
        [
            "DOC004", "A002", None, "Certificat Profesional", "CERT-IT-2024",
            date(2024, 6, 1), date(2026, 6, 1), "IT Academy",
            None, "Certificare Cloud AWS",
        ],
        [
            "DOC005", "A003", None, "Certificat Medical", "CM-2025-001",
            date(2025, 1, 15), date(2026, 1, 15), "Clinica Medicală X",
            None, "Apt pentru muncă",
        ],
        [
            "DOC006", "A003", None, "Cazier Judiciar", "CJ-2025-001",
            date(2025, 1, 10), date(2025, 7, 10), "Poliția Sector 1",
            None, "",
        ],
        [
            "DOC007", "A004", None, "Diplomă Studii", "DIP-2010-001",
            date(2010, 7, 15), None, "ASE București",
            None, "Licență Contabilitate",
        ],
        [
            "DOC008", "A005", None, "Permis Conducere", "PC-B-12345",
            date(2018, 3, 20), date(2033, 3, 20), "DRPCIV",
            None, "Categoria B",
        ],
    ]

    ws = create_sheet_with_headers(
        wb, "Documente", "DOCUMENTE ANGAJAȚI", headers, demo_data
    )

    # Formule VLOOKUP pentru Nume Angajat
    for row in range(3, 3 + len(demo_data)):
        ws.cell(row=row, column=3).value = (
            f'=IFERROR(VLOOKUP(B{row},\'Angajați\'!$A:$C,2,0)&" "'
            f'&VLOOKUP(B{row},\'Angajați\'!$A:$C,3,0),"N/A")'
        )

    # Formula Status automat
    for row in range(3, 3 + len(demo_data)):
        ws.cell(row=row, column=9).value = (
            f'=IF(G{row}="","Valid",'
            f'IF(G{row}<TODAY(),"Expirat",'
            f'IF(G{row}<=TODAY()+30,"Expiră Curând","Valid")))'
        )

    # Formatare date
    for col in [6, 7]:
        for row in ws.iter_rows(min_row=3, max_row=ws.max_row,
                                 min_col=col, max_col=col):
            for cell in row:
                if cell.value:
                    cell.number_format = NUMBER_FORMAT_DATE

    # Validări
    add_dropdown_validation(ws, "D3:D1000", TIPURI_DOCUMENT)

    # Formatare condițională - Status
    add_status_conditional_formatting(
        ws, "I", 3, 1000,
        {
            "Valid": COLORS["light_green"],
            "Expirat": COLORS["light_red"],
            "Expiră Curând": COLORS["light_yellow"],
            "Lipsă": COLORS["light_orange"],
        }
    )

    # Formatare condițională - Data Expirării
    add_date_expiry_formatting(ws, "G", 3, 1000, days_warning=30)

    return ws
