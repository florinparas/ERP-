"""
Foaia Recrutare - Proces recrutare candidați
"""
from datetime import date

from config.hr_config import (
    STATUS_RECRUTARE, SURSE_RECRUTARE,
    NUMBER_FORMAT_DATE, NUMBER_FORMAT_RON, COLORS
)
from utils.helpers import (
    create_sheet_with_headers, add_dropdown_validation,
    add_status_conditional_formatting, add_number_validation
)


def create_sheet(wb):
    """Creează foaia Recrutare"""
    headers = [
        ("ID", 8),
        ("Post", 22),
        ("Departament", 18),
        ("Data Deschidere", 14),
        ("Nume Candidat", 20),
        ("Email", 25),
        ("Telefon", 14),
        ("Sursă", 14),
        ("Data Aplicării", 14),
        ("Status", 16),
        ("Data Interviu", 14),
        ("Interviewer", 18),
        ("Scor (1-10)", 10),
        ("Ofertă Salarială (RON)", 18),
        ("Observații", 25),
    ]

    demo_data = [
        [
            "R001", "Developer Software Mid", "IT & Dezvoltare",
            date(2025, 2, 1), "Popa Alexandru", "alex.popa@email.com",
            "0731000010", "LinkedIn", date(2025, 2, 5),
            "Interviu Tehnic", date(2025, 3, 10), "Ionescu Maria",
            8, 9000, "Experiență 3 ani Java/Python",
        ],
        [
            "R002", "Developer Software Mid", "IT & Dezvoltare",
            date(2025, 2, 1), "Stan Mihaela", "mihaela.stan@email.com",
            "0731000011", "eJobs", date(2025, 2, 8),
            "Screening CV", None, "",
            None, None, "CV promițător, de programat interviu",
        ],
        [
            "R003", "Specialist Marketing", "Marketing",
            date(2025, 3, 1), "Radu Cristina", "cristina.radu@email.com",
            "0731000012", "Referral", date(2025, 3, 5),
            "Ofertă", date(2025, 3, 15), "Constantinescu Dan",
            9, 7000, "Referral de la echipa curentă",
        ],
        [
            "R004", "Operator Logistică", "Logistică",
            date(2025, 3, 1), "Nistor Bogdan", "bogdan.nistor@email.com",
            "0731000013", "BestJobs", date(2025, 3, 3),
            "Respins", date(2025, 3, 8), "Vasilescu Mihai",
            3, None, "Nu corespunde cerințelor",
        ],
    ]

    ws = create_sheet_with_headers(
        wb, "Recrutare", "MANAGEMENT RECRUTARE", headers, demo_data
    )

    # Formatare date
    for col in [4, 9, 11]:
        for row in ws.iter_rows(min_row=3, max_row=ws.max_row,
                                 min_col=col, max_col=col):
            for cell in row:
                if cell.value:
                    cell.number_format = NUMBER_FORMAT_DATE

    # Formatare salariu
    for row in ws.iter_rows(min_row=3, max_row=ws.max_row, min_col=14, max_col=14):
        for cell in row:
            if cell.value:
                cell.number_format = NUMBER_FORMAT_RON

    # Validări
    add_dropdown_validation(ws, "C3:C1000", "=Departamente!$B$3:$B$50")
    add_dropdown_validation(ws, "H3:H1000", SURSE_RECRUTARE)
    add_dropdown_validation(ws, "J3:J1000", STATUS_RECRUTARE)
    add_number_validation(ws, "M3:M1000", 1, 10)

    # Formatare condițională - Status recrutare
    add_status_conditional_formatting(
        ws, "J", 3, 1000,
        {
            "Angajat": COLORS["light_green"],
            "Respins": COLORS["light_red"],
            "Retras": COLORS["light_red"],
            "Ofertă": COLORS["light_blue"],
            "Interviu Tehnic": COLORS["light_yellow"],
            "Interviu Final": COLORS["light_yellow"],
            "Screening CV": COLORS["light_orange"],
            "Nou": "FFFFFF",
        }
    )

    return ws
