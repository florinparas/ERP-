# ERP System cu Integrare Microsoft Office

Sistem ERP complet pentru gestionarea afacerii, cu integrare nativa cu suita Microsoft Office (Excel, Word, Outlook).

## Module

- **Dashboard** - Vizualizare de ansamblu cu statistici in timp real
- **Clienti** - Gestionare baza de date clienti (CUI, adresa, contact)
- **Produse** - Catalog produse cu preturi, TVA, categorii
- **Facturi** - Creare, vizualizare si gestionare facturi cu calcul automat TVA
- **Comenzi** - Gestionare comenzi cu urmarire status
- **Angajati** - Evidenta angajati (pozitie, departament, salariu)
- **Stocuri** - Gestiune stocuri cu alerte stoc minim

## Integrare Microsoft Office

### Excel (openpyxl)
- Export date in format `.xlsx` cu formatare profesionala (stiluri, culori, auto-width)
- Import clienti si produse din fisiere Excel
- Rapoarte cu foi multiple si sumar

### Word (python-docx)
- Generare facturi in format `.docx` cu tabele formatate
- Generare comenzi cu detalii complete
- Rapoarte generale cu statistici

### Outlook / Email (SMTP Office 365)
- Trimitere facturi pe email cu atasament Word
- Confirmare comenzi pe email
- Template-uri HTML profesionale

## Instalare

```bash
pip install -r requirements.txt
python app.py
```

Acceseaza: http://localhost:5000
Login: admin / admin123

## Configurare Email (Outlook)

Seteaza variabilele de mediu:
```
MAIL_SERVER=smtp.office365.com
MAIL_PORT=587
MAIL_USERNAME=email@company.com
MAIL_PASSWORD=parola
MAIL_DEFAULT_SENDER=email@company.com
```

## Structura Proiect

```
ERP-/
├── app.py                    # Punct de intrare
├── config.py                 # Configurare
├── requirements.txt          # Dependente
├── erp_app/
│   ├── __init__.py          # Factory Flask
│   ├── models/              # Modele baza de date
│   │   ├── user.py          # Utilizatori si autentificare
│   │   ├── client.py        # Clienti
│   │   ├── product.py       # Produse
│   │   ├── invoice.py       # Facturi + articole
│   │   ├── order.py         # Comenzi + articole
│   │   ├── employee.py      # Angajati
│   │   └── inventory.py     # Miscari stoc
│   ├── routes/              # Rute web
│   │   ├── auth.py          # Login/Register
│   │   ├── dashboard.py     # Pagina principala
│   │   ├── clients.py       # CRUD Clienti
│   │   ├── products.py      # CRUD Produse
│   │   ├── invoices.py      # CRUD Facturi
│   │   ├── orders.py        # CRUD Comenzi
│   │   ├── employees.py     # CRUD Angajati
│   │   ├── inventory.py     # Gestiune stoc
│   │   └── office.py        # Export/Import Office
│   ├── services/            # Servicii integrare Office
│   │   ├── excel_service.py # Export/Import Excel
│   │   ├── word_service.py  # Generare documente Word
│   │   └── email_service.py # Trimitere email Outlook
│   ├── templates/           # Template-uri HTML
│   └── static/              # CSS, JS, fisiere exportate
```

## Tehnologii

- **Backend:** Python, Flask, SQLAlchemy, SQLite
- **Office:** openpyxl (Excel), python-docx (Word), smtplib (Outlook)
- **Frontend:** HTML5, CSS3, JavaScript vanilla
