# ERP Management - Modul HR

Sistem ERP în Excel pentru managementul de personal, generat cu Python (openpyxl).

## Modul HR - Management Personal

Fișierul Excel generat conține 11 foi:

| Foaie | Descriere |
|-------|-----------|
| Dashboard | Panou principal cu KPI-uri, grafice și navigare rapidă |
| Angajați | Evidența completă a personalului |
| Contracte | Contracte de muncă cu VLOOKUP la angajați |
| Departamente | Departamente & funcții cu grile salariale |
| Pontaj | Pontaj lunar cu coduri prezență și totaluri automate |
| Concedii | Management concedii cu calcul zile rămase |
| Salarizare | Calcul salarii conform legislației RO 2025 (CAS 25%, CASS 10%, Impozit 10%) |
| Evaluări | Evaluări performanță cu scoruri 1-5 |
| Training | Cursuri & formări profesionale |
| Recrutare | Proces recrutare candidați |
| Configurare | Parametri fiscali, date companie, liste dropdown |

## Instalare & Generare

```bash
pip install -r requirements.txt
python generate_hr.py
```

Fișierul `ERP_HR_Module.xlsx` va fi generat în directorul curent.

## Module VBA (opțional)

Directorul `vba/` conține module VBA care pot fi importate în Excel:
- `Module_Navigation.bas` - Navigare rapidă între foi
- `Module_Pontaj.bas` - Generare automată pontaj lunar
- `Module_Salarizare.bas` - Calcul automat salarizare & fluturași
- `Module_Rapoarte.bas` - Rapoarte (departamente, concedii, costuri)
- `Module_Utils.bas` - Validare CNP, backup, protecție foi

Pentru import: Excel → Alt+F11 → File → Import File → selectați fișierele .bas

## Parametri Fiscali România 2025

- CAS (Pensie angajat): 25%
- CASS (Sănătate angajat): 10%
- Impozit pe venit: 10% (flat tax)
- CAM (Angajator): 2.25%
- Salariu minim brut: 4.050 RON
