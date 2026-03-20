"""
Foaia Salarizare - Calcul salarii conform legislației RO 2025
"""
from config.hr_config import (
    CAS_RATE, CASS_RATE, TAX_RATE, CAM_RATE, SALARIU_MINIM_BRUT,
    NUMBER_FORMAT_RON, NUMBER_FORMAT_PERCENT, NUMBER_FORMAT_DATE,
    NUMBER_FORMAT_INT, COLORS
)
from utils.helpers import (
    create_sheet_with_headers, add_status_conditional_formatting
)
from utils.styles import get_border
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import PatternFill, Font


def create_sheet(wb):
    """Creează foaia Salarizare cu formule calcul conform legislației RO"""
    headers = [
        ("ID", 8),
        ("ID Angajat", 12),
        ("Nume Angajat", 20),
        ("Luna", 8),
        ("An", 8),
        ("Salariu Brut", 16),
        ("Zile Lucrate", 12),
        ("Zile Totale Lună", 14),
        ("Salariu Proporțional", 18),
        ("CAS 25%", 14),
        ("CASS 10%", 14),
        ("Baza Impozabilă", 16),
        ("Deducere Personală", 16),
        ("Impozit 10%", 14),
        ("Alte Deduceri", 14),
        ("Tichete Masă", 14),
        ("Salariu Net", 16),
        ("CAM 2.25% (angajator)", 18),
        ("Cost Total Angajator", 18),
    ]

    # Date demo - luna martie 2025
    demo_employees = [
        ("S001", "A001", 3, 2025, 25000),
        ("S002", "A002", 3, 2025, 14000),
        ("S003", "A003", 3, 2025, 7500),
        ("S004", "A004", 3, 2025, 9000),
        ("S005", "A005", 3, 2025, 12000),
    ]

    # Creăm foaia cu headere doar (datele le populăm cu formule)
    ws = create_sheet_with_headers(
        wb, "Salarizare", "STAT DE PLATĂ - SALARIZARE", headers
    )

    for row_idx, (sal_id, emp_id, luna, an, brut) in enumerate(demo_employees, 3):
        r = row_idx

        # ID
        ws.cell(row=r, column=1, value=sal_id).border = get_border()

        # ID Angajat
        ws.cell(row=r, column=2, value=emp_id).border = get_border()

        # Nume Angajat (VLOOKUP)
        ws.cell(row=r, column=3).value = (
            f'=IFERROR(VLOOKUP(B{r},\'Angajați\'!$A:$C,2,0)&" "'
            f'&VLOOKUP(B{r},\'Angajați\'!$A:$C,3,0),"N/A")'
        )
        ws.cell(row=r, column=3).border = get_border()

        # Luna, An
        ws.cell(row=r, column=4, value=luna).border = get_border()
        ws.cell(row=r, column=5, value=an).border = get_border()

        # Salariu Brut
        cell_brut = ws.cell(row=r, column=6, value=brut)
        cell_brut.number_format = NUMBER_FORMAT_RON
        cell_brut.border = get_border()

        # Zile Lucrate (din pontaj)
        ws.cell(row=r, column=7).value = (
            f'=IFERROR(SUMPRODUCT((Pontaj!$A$3:$A$200=B{r})'
            f'*(Pontaj!$D$3:$D$200=D{r})'
            f'*(Pontaj!$E$3:$E$200=E{r})'
            f'*Pontaj!$AK$3:$AK$200),0)'
        )
        ws.cell(row=r, column=7).border = get_border()

        # Zile Totale Lună (zile lucrătoare în luna respectivă)
        ws.cell(row=r, column=8).value = (
            f'=NETWORKDAYS(DATE(E{r},D{r},1),'
            f'EOMONTH(DATE(E{r},D{r},1),0))'
        )
        ws.cell(row=r, column=8).border = get_border()

        # Salariu Proporțional = Brut * Zile Lucrate / Zile Totale
        ws.cell(row=r, column=9).value = (
            f'=IF(H{r}>0,ROUND(F{r}*G{r}/H{r},2),0)'
        )
        ws.cell(row=r, column=9).number_format = NUMBER_FORMAT_RON
        ws.cell(row=r, column=9).border = get_border()

        # CAS 25% (din salariul proporțional)
        ws.cell(row=r, column=10).value = f'=ROUND(I{r}*{CAS_RATE},2)'
        ws.cell(row=r, column=10).number_format = NUMBER_FORMAT_RON
        ws.cell(row=r, column=10).border = get_border()

        # CASS 10%
        ws.cell(row=r, column=11).value = f'=ROUND(I{r}*{CASS_RATE},2)'
        ws.cell(row=r, column=11).number_format = NUMBER_FORMAT_RON
        ws.cell(row=r, column=11).border = get_border()

        # Baza Impozabilă = Salariu Proporțional - CAS - CASS - Deducere Personală
        ws.cell(row=r, column=12).value = (
            f'=MAX(I{r}-J{r}-K{r}-M{r},0)'
        )
        ws.cell(row=r, column=12).number_format = NUMBER_FORMAT_RON
        ws.cell(row=r, column=12).border = get_border()

        # Deducere Personală (conform legislație - simplificat)
        # Se aplică doar pentru salarii sub un anumit prag
        ws.cell(row=r, column=13).value = (
            f'=IF(I{r}<=Configurare!$B$7*2,Configurare!$B$8,0)'
        )
        ws.cell(row=r, column=13).number_format = NUMBER_FORMAT_RON
        ws.cell(row=r, column=13).border = get_border()

        # Impozit 10%
        ws.cell(row=r, column=14).value = f'=ROUND(L{r}*{TAX_RATE},2)'
        ws.cell(row=r, column=14).number_format = NUMBER_FORMAT_RON
        ws.cell(row=r, column=14).border = get_border()

        # Alte Deduceri (manual)
        ws.cell(row=r, column=15, value=0).border = get_border()
        ws.cell(row=r, column=15).number_format = NUMBER_FORMAT_RON

        # Tichete Masă (manual)
        ws.cell(row=r, column=16, value=0).border = get_border()
        ws.cell(row=r, column=16).number_format = NUMBER_FORMAT_RON

        # Salariu Net = Proporțional - CAS - CASS - Impozit - Alte Deduceri + Tichete
        ws.cell(row=r, column=17).value = (
            f'=I{r}-J{r}-K{r}-N{r}-O{r}+P{r}'
        )
        ws.cell(row=r, column=17).number_format = NUMBER_FORMAT_RON
        ws.cell(row=r, column=17).border = get_border()

        # CAM 2.25% (contribuție angajator)
        ws.cell(row=r, column=18).value = f'=ROUND(I{r}*{CAM_RATE},2)'
        ws.cell(row=r, column=18).number_format = NUMBER_FORMAT_RON
        ws.cell(row=r, column=18).border = get_border()

        # Cost Total Angajator = Salariu Proporțional + CAM
        ws.cell(row=r, column=19).value = f'=I{r}+R{r}'
        ws.cell(row=r, column=19).number_format = NUMBER_FORMAT_RON
        ws.cell(row=r, column=19).border = get_border()

    # Rând TOTAL
    total_row = 3 + len(demo_employees)
    total_font = Font(name="Calibri", size=11, bold=True)
    ws.cell(row=total_row, column=1, value="TOTAL").font = total_font
    ws.cell(row=total_row, column=1).border = get_border()

    # Totaluri pe coloanele numerice
    for col in [6, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19]:
        from openpyxl.utils import get_column_letter
        col_letter = get_column_letter(col)
        ws.cell(row=total_row, column=col).value = (
            f'=SUM({col_letter}3:{col_letter}{total_row - 1})'
        )
        ws.cell(row=total_row, column=col).number_format = NUMBER_FORMAT_RON
        ws.cell(row=total_row, column=col).font = total_font
        ws.cell(row=total_row, column=col).border = get_border()

    # Formatare condițională - Salariu Net sub minim
    net_range = f"Q3:Q1000"
    rule = FormulaRule(
        formula=[f'AND(Q3<>"",Q3<{SALARIU_MINIM_BRUT}*0.585)'],
        fill=PatternFill(start_color=COLORS["light_red"],
                        end_color=COLORS["light_red"], fill_type="solid"),
        font=Font(color="FF0000", bold=True)
    )
    ws.conditional_formatting.add(net_range, rule)

    # Adaugă legendă formule
    legend_row = total_row + 2
    ws.cell(row=legend_row, column=1, value="FORMULE CALCUL:").font = total_font
    formulas_info = [
        "CAS = Salariu Proporțional × 25% (contribuție pensie)",
        "CASS = Salariu Proporțional × 10% (contribuție sănătate)",
        "Baza Impozabilă = Salariu Proporțional - CAS - CASS - Deducere Personală",
        "Impozit = Baza Impozabilă × 10%",
        "Salariu Net = Salariu Proporțional - CAS - CASS - Impozit - Alte Deduceri + Tichete Masă",
        "CAM = Salariu Proporțional × 2.25% (contribuție angajator)",
        "Cost Total Angajator = Salariu Proporțional + CAM",
    ]
    for i, formula_text in enumerate(formulas_info, legend_row + 1):
        ws.cell(row=i, column=1, value=formula_text).font = Font(
            name="Calibri", size=9, italic=True, color="666666"
        )

    return ws
