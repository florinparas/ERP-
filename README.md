# ERP Management - Modul HR

Sistem ERP în Excel pentru managementul de personal, generat cu Python (openpyxl).

## Modul HR - Management Personal

Fișierul Excel generat conține **16 foi**:

| Foaie | Descriere |
|-------|-----------|
| Dashboard | Panou principal cu KPI-uri, grafic distribuție, alerte, navigare rapidă |
| Angajați | Evidența completă a personalului cu validări și formatare condițională |
| Contracte | Contracte de muncă cu VLOOKUP, alertă expirare |
| Departamente | Departamente & funcții cu grile salariale, COUNTIF angajați |
| Documente | Management documente angajați cu status automat (valid/expirat) |
| Pontaj | Pontaj lunar cu 31 coloane zile, coduri colorate, totaluri automate |
| Ore Suplimentare | Evidență ore suplimentare cu calcul valoare brută (spor 75%/100%) |
| Concedii | Management concedii cu calcul zile rămase (NETWORKDAYS) |
| Salarizare | Stat de plată complet: brut, ore suplimentare, deducere personală, net |
| Fluturași | Fluturași de salariu print-ready (selectare angajat dinamic) |
| Evaluări | Evaluări performanță cu scoruri 1-5 și scor final automat |
| Training | Cursuri & formări profesionale cu alertă expirare certificări |
| Recrutare | Proces recrutare candidați cu pipeline statusuri |
| Organigramă | Vizualizare ierarhie organizațională (date + tree view) |
| Istoric | Audit trail - jurnal modificări (manual sau VBA) |
| Configurare | Parametri fiscali, date companie, sărbători legale, liste dropdown |

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
- Deducere personală: variabilă (funcție de venit și persoane în întreținere)
