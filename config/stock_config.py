"""
Configurare StockAgent — Constante, culori, reguli risk, date demo
Aplicație Management Portofoliu Investiții
"""

# ============================================================================
# CULORI TEMĂ TRADING (hex fără #, pentru openpyxl)
# Aceleași chei ca HR COLORS pentru compatibilitate cu utils/styles.py
# ============================================================================
COLORS = {
    "header_bg": "1B2A4A",       # Albastru închis (trading dark)
    "header_font": "FFFFFF",     # Alb
    "row_alt": "E8EDF5",         # Albastru-gri deschis (zebra)
    "row_normal": "FFFFFF",      # Alb
    "accent": "2E75B6",          # Albastru accent
    "success": "27AE60",         # Verde profit
    "warning": "F39C12",         # Galben/Portocaliu
    "danger": "E74C3C",          # Roșu pierdere
    "info": "3498DB",            # Albastru info
    "light_green": "D5F5E3",     # Verde deschis (profit)
    "light_red": "FADBD8",       # Roșu deschis (pierdere)
    "light_yellow": "FEF9E7",    # Galben deschis (atenție)
    "light_blue": "D6EAF8",     # Albastru deschis
    "light_orange": "FDEBD0",    # Portocaliu deschis
    "border": "AEB6BF",          # Gri bordură
    "title_bg": "0B1929",        # Foarte închis (titluri)
    # Culori specifice trading
    "profit": "27AE60",          # Verde profit
    "loss": "E74C3C",            # Roșu pierdere
    "neutral": "F39C12",         # Galben neutru
    "gold": "F1C40F",            # Aur (watchlist)
    "purple": "8E44AD",          # Violet (analiză)
    "teal": "1ABC9C",            # Teal (fundamental)
}

# ============================================================================
# REGULI RISK MANAGEMENT (hardcoded — fără excepții)
# ============================================================================
MAX_POSITION_PCT = 0.10                  # 10% max per poziție standard
MAX_POSITION_SPECULATIVE_PCT = 0.05      # 5% max per poziție speculativă
MAX_SECTOR_PCT = 0.30                    # 30% max per sector
MIN_CASH_RESERVE_PCT = 0.10              # 10% cash minim permanent
STOP_LOSS_DEFAULT_PCT = 0.07             # 7% stop-loss standard
STOP_LOSS_MAX_PCT = 0.10                 # 10% stop-loss maxim
TRAILING_STOP_BREAKEVEN_TRIGGER = 0.15   # Mută SL la breakeven după +15%
TRAILING_STOP_TRAIL_TRIGGER = 0.25       # -10% trailing după +25%
TRAILING_STOP_TRAIL_PCT = 0.10           # Trail 10% de la maxim
MIN_RR_RATIO = 2.0                       # Risk/Reward minim standard
MIN_RR_SPECULATIVE = 1.5                 # R/R minim speculativ (setup A+)
DRAWDOWN_REDUCE_PCT = 0.10              # Reduce expunere la -10% drawdown
DRAWDOWN_EXIT_PCT = 0.15                # Ieșire pe cash la -15% drawdown
MAX_CONSECUTIVE_LOSSES = 3               # Pauză 48h după 3 SL consecutive

# ============================================================================
# LISTE DROPDOWN
# ============================================================================
SECTOARE = [
    "Energie", "Financiar", "Tehnologie", "Sănătate", "Consum Discretionar",
    "Consum de Bază", "Industrie", "Materiale", "Utilități", "Imobiliare",
    "Comunicații", "Servicii Petroliere", "Minerit Aur", "Shipping",
]

TIP_TRANZACTIE = ["Cumpărare", "Vânzare", "Vânzare Short", "Acoperire Short"]

TIP_POZITIE = ["Standard", "Speculativă"]

STATUS_POZITIE = ["Deschisă", "Închisă", "Parțial Închisă"]

CONVINGERE = ["★", "★★", "★★★", "★★★★", "★★★★★"]

CONVINGERE_NUMERIC = [1, 2, 3, 4, 5]

TIMEFRAME = [
    "Swing (1-4 săpt.)",
    "Tactică (1-3 luni)",
    "Speculativă (1-5 zile)",
]

PIATA = ["BVB", "NYSE", "NASDAQ", "XETRA", "LSE", "Euronext"]

MONEDA = ["RON", "USD", "EUR", "GBP"]

TREND = ["Bullish", "Bearish", "Neutru", "Consolidare"]

SEMNAL = [
    "Cumpărare Puternică",
    "Cumpărare",
    "Neutru",
    "Vânzare",
    "Vânzare Puternică",
]

VERDICT_FUNDAMENTAL = ["BUY", "HOLD", "SELL", "WATCH"]

# ============================================================================
# FORMATE NUMERICE
# ============================================================================
NUMBER_FORMAT_CURRENCY = '#,##0.00'
NUMBER_FORMAT_USD = '#,##0.00 "$"'
NUMBER_FORMAT_RON = '#,##0.00 "RON"'
NUMBER_FORMAT_EUR = '#,##0.00 "€"'
NUMBER_FORMAT_PERCENT = '0.00%'
NUMBER_FORMAT_PRICE = '#,##0.00'
NUMBER_FORMAT_DATE = 'DD.MM.YYYY'
NUMBER_FORMAT_INT = '#,##0'
NUMBER_FORMAT_RR = '0.0\\:1'

# ============================================================================
# DATE DEMO — PORTOFOLIU STOCKAGENT
# ============================================================================
PORTFOLIO_INITIAL_CAPITAL = 500000   # RON
PORTFOLIO_CASH = 127500              # RON (25.5% — rezervă agresivă dar disciplinată)
PORTFOLIO_START_DATE = "2025-01-02"
PORTFOLIO_PEAK_VALUE = 520000        # High watermark

POZITII_DEMO = [
    {
        "id": "P001", "simbol": "TLV", "denumire": "Banca Transilvania",
        "piata": "BVB", "sector": "Financiar", "tip": "Standard",
        "cantitate": 5000, "pret_intrare": 28.50, "data_intrare": "2025-01-15",
        "pret_curent": 31.80, "stop_loss": 26.50, "target1": 34.00,
        "target2": 38.00, "moneda": "RON",
    },
    {
        "id": "P002", "simbol": "SNG", "denumire": "Romgaz",
        "piata": "BVB", "sector": "Energie", "tip": "Standard",
        "cantitate": 2000, "pret_intrare": 52.00, "data_intrare": "2025-02-03",
        "pret_curent": 56.40, "stop_loss": 48.00, "target1": 60.00,
        "target2": 66.00, "moneda": "RON",
    },
    {
        "id": "P003", "simbol": "SNP", "denumire": "OMV Petrom",
        "piata": "BVB", "sector": "Energie", "tip": "Standard",
        "cantitate": 80000, "pret_intrare": 0.625, "data_intrare": "2025-01-20",
        "pret_curent": 0.598, "stop_loss": 0.565, "target1": 0.700,
        "target2": 0.780, "moneda": "RON",
    },
    {
        "id": "P004", "simbol": "FP", "denumire": "Fondul Proprietatea",
        "piata": "BVB", "sector": "Financiar", "tip": "Standard",
        "cantitate": 25000, "pret_intrare": 2.15, "data_intrare": "2025-02-10",
        "pret_curent": 2.32, "stop_loss": 1.98, "target1": 2.50,
        "target2": 2.80, "moneda": "RON",
    },
    {
        "id": "P005", "simbol": "DIGI", "denumire": "Digi Communications",
        "piata": "BVB", "sector": "Comunicații", "tip": "Standard",
        "cantitate": 200, "pret_intrare": 45.00, "data_intrare": "2025-03-01",
        "pret_curent": 48.50, "stop_loss": 41.50, "target1": 52.00,
        "target2": 58.00, "moneda": "RON",
    },
    {
        "id": "P006", "simbol": "BRD", "denumire": "BRD Groupe Société Générale",
        "piata": "BVB", "sector": "Financiar", "tip": "Speculativă",
        "cantitate": 3000, "pret_intrare": 19.20, "data_intrare": "2025-03-10",
        "pret_curent": 18.60, "stop_loss": 17.80, "target1": 21.50,
        "target2": 23.00, "moneda": "RON",
    },
]

TRANZACTII_DEMO = [
    {
        "id": "T001", "data_intrare": "2025-01-08", "simbol": "H2O",
        "denumire": "H2O Innovation", "tip": "Cumpărare",
        "cantitate": 10000, "pret_intrare": 0.42, "data_iesire": "2025-01-22",
        "pret_iesire": 0.51, "comision": 42, "motiv_intrare": "Breakout volum peste MA50",
        "motiv_iesire": "Target 1 atins", "lectii": "Setup clasic, executat corect",
    },
    {
        "id": "T002", "data_intrare": "2025-01-10", "simbol": "SNN",
        "denumire": "Nuclearelectrica", "tip": "Cumpărare",
        "cantitate": 500, "pret_intrare": 52.00, "data_iesire": "2025-02-05",
        "pret_iesire": 48.20, "comision": 26, "motiv_intrare": "Support bounce la MA200",
        "motiv_iesire": "Stop-loss activat", "lectii": "Trend primar bearish, nu intri contra trendului",
    },
    {
        "id": "T003", "data_intrare": "2025-02-12", "simbol": "TEL",
        "denumire": "Transelectrica", "tip": "Cumpărare",
        "cantitate": 1500, "pret_intrare": 30.00, "data_iesire": "2025-03-05",
        "pret_iesire": 34.50, "comision": 45, "motiv_intrare": "Cup & Handle pe weekly",
        "motiv_iesire": "Target 2 atins", "lectii": "Rabdare recompensată, pattern weekly de încredere",
    },
    {
        "id": "T004", "data_intrare": "2025-02-20", "simbol": "WINE",
        "denumire": "Purcari Wineries", "tip": "Cumpărare",
        "cantitate": 800, "pret_intrare": 18.50, "data_iesire": "2025-03-01",
        "pret_iesire": 17.20, "comision": 15, "motiv_intrare": "Ascending triangle",
        "motiv_iesire": "Stop-loss activat — false breakout",
        "lectii": "Volumul la breakout era sub medie. Confirmare volum obligatorie.",
    },
    {
        "id": "T005", "data_intrare": "2025-03-05", "simbol": "ONE",
        "denumire": "One United Properties", "tip": "Cumpărare",
        "cantitate": 5000, "pret_intrare": 1.05, "data_iesire": "2025-03-15",
        "pret_iesire": 1.18, "comision": 10, "motiv_intrare": "Golden Cross + volum crescător",
        "motiv_iesire": "Target 1 atins, trailing stop pe rest",
        "lectii": "Entry excelent pe confluență indicatori. De repetat.",
    },
]

WATCHLIST_DEMO = [
    {
        "simbol": "SNN", "denumire": "Nuclearelectrica", "sector": "Energie",
        "piata": "BVB", "pret_curent": 48.50, "pret_tinta": 46.00,
        "stop_loss": 42.50, "target": 55.00, "convingere": 4,
        "timeframe": "Swing (1-4 săpt.)",
        "catalizator": "Rezultate Q1 2025 — 28 aprilie",
        "observatii": "Retestare suport solid la 46. RSI la 38, aproape de supravândut. "
                      "Aștept confirmare pe volum crescător la bounce.",
    },
    {
        "simbol": "M", "denumire": "Medlife", "sector": "Sănătate",
        "piata": "BVB", "pret_curent": 12.80, "pret_tinta": 12.00,
        "stop_loss": 11.00, "target": 15.00, "convingere": 3,
        "timeframe": "Tactică (1-3 luni)",
        "catalizator": "Achiziție nouă clinică — aprobare CNA",
        "observatii": "Consolidare strânsă 3 săptămâni. Bollinger squeeze. "
                      "Potențial breakout fie sus fie jos. Aștept direcție.",
    },
    {
        "simbol": "COTE", "denumire": "Conpet", "sector": "Energie",
        "piata": "BVB", "pret_curent": 88.00, "pret_tinta": 85.00,
        "stop_loss": 78.00, "target": 100.00, "convingere": 4,
        "timeframe": "Swing (1-4 săpt.)",
        "catalizator": "Dividend yield 8%+ — ex-date aproape",
        "observatii": "Dividend play clasic. Prețul subestimează yield-ul. "
                      "Intrare la pullback spre 85, SL sub suport major.",
    },
    {
        "simbol": "AQ", "denumire": "Aquila Part Prod", "sector": "Consum de Bază",
        "piata": "BVB", "pret_curent": 1.35, "pret_tinta": 1.28,
        "stop_loss": 1.18, "target": 1.55, "convingere": 3,
        "timeframe": "Tactică (1-3 luni)",
        "catalizator": "Expansiune retail — contract Kaufland",
        "observatii": "Bull flag pe daily. MACD bullish cross iminent. "
                      "Volum mediu OK dar nu spectaculos. Sizing conservator.",
    },
    {
        "simbol": "TRP", "denumire": "Teraplast", "sector": "Materiale",
        "piata": "BVB", "pret_curent": 0.68, "pret_tinta": 0.62,
        "stop_loss": 0.56, "target": 0.80, "convingere": 5,
        "timeframe": "Swing (1-4 săpt.)",
        "catalizator": "Contracte infrastructură PNRR — anunț martie",
        "observatii": "Setup A+ — cup and handle pe weekly, breakout iminent. "
                      "Volum explodat +180% peste medie. Conviction maximă. "
                      "Dacă trece de 0.70, accelerare puternică.",
    },
]

# ============================================================================
# ANALIZĂ TEHNICĂ — DATE DEMO
# ============================================================================
ANALIZA_TEHNICA_DEMO = [
    {
        "simbol": "TLV", "pret": 31.80, "ma50": 30.20, "ma200": 27.50,
        "trend_ma": "Bullish", "cross": "Golden Cross",
        "rsi": 62, "semnal_rsi": "Neutru", "macd": 0.85, "semnal_macd": "Bullish",
        "stoch_k": 71, "stoch_d": 65, "semnal_stoch": "Neutru",
        "volum_mediu": 1200000, "volum_curent": 1450000, "obv_trend": "Crescător",
        "suport1": 30.00, "suport2": 28.00, "rezistenta1": 33.00, "rezistenta2": 35.50,
        "pattern": "Bull Flag", "pattern_tf": "Daily", "pattern_dir": "Bullish",
    },
    {
        "simbol": "SNG", "pret": 56.40, "ma50": 54.00, "ma200": 50.00,
        "trend_ma": "Bullish", "cross": "-",
        "rsi": 58, "semnal_rsi": "Neutru", "macd": 1.20, "semnal_macd": "Bullish",
        "stoch_k": 55, "stoch_d": 52, "semnal_stoch": "Neutru",
        "volum_mediu": 350000, "volum_curent": 420000, "obv_trend": "Crescător",
        "suport1": 54.00, "suport2": 51.00, "rezistenta1": 58.50, "rezistenta2": 62.00,
        "pattern": "Ascending Triangle", "pattern_tf": "Weekly", "pattern_dir": "Bullish",
    },
    {
        "simbol": "SNP", "pret": 0.598, "ma50": 0.615, "ma200": 0.580,
        "trend_ma": "Neutru", "cross": "-",
        "rsi": 42, "semnal_rsi": "Neutru", "macd": -0.008, "semnal_macd": "Bearish",
        "stoch_k": 35, "stoch_d": 40, "semnal_stoch": "Bearish",
        "volum_mediu": 8500000, "volum_curent": 6200000, "obv_trend": "Scăzător",
        "suport1": 0.580, "suport2": 0.550, "rezistenta1": 0.625, "rezistenta2": 0.660,
        "pattern": "Consolidare", "pattern_tf": "Daily", "pattern_dir": "Neutru",
    },
]

# ============================================================================
# ANALIZĂ FUNDAMENTALĂ — DATE DEMO
# ============================================================================
ANALIZA_FUNDAMENTALA_DEMO = [
    {
        "simbol": "TLV", "denumire": "Banca Transilvania", "sector": "Financiar",
        "pret": 31.80, "capitalizare": 18500, "revenue_growth": 18.5,
        "eps_growth": 22.3, "pe": 6.8, "pe_sector": 8.5, "ev_ebitda": 5.2,
        "fcf_yield": 8.5, "roe": 18.2, "debt_equity": 0.9, "div_yield": 5.2,
        "payout": 35, "scor": 9,
    },
    {
        "simbol": "SNG", "denumire": "Romgaz", "sector": "Energie",
        "pret": 56.40, "capitalizare": 21800, "revenue_growth": 12.0,
        "eps_growth": 15.8, "pe": 5.2, "pe_sector": 7.0, "ev_ebitda": 3.8,
        "fcf_yield": 12.0, "roe": 22.5, "debt_equity": 0.3, "div_yield": 7.8,
        "payout": 40, "scor": 8,
    },
    {
        "simbol": "SNP", "denumire": "OMV Petrom", "sector": "Energie",
        "pret": 0.598, "capitalizare": 33900, "revenue_growth": 8.5,
        "eps_growth": -5.2, "pe": 7.1, "pe_sector": 7.0, "ev_ebitda": 4.5,
        "fcf_yield": 9.2, "roe": 15.0, "debt_equity": 0.4, "div_yield": 6.5,
        "payout": 45, "scor": 7,
    },
    {
        "simbol": "FP", "denumire": "Fondul Proprietatea", "sector": "Financiar",
        "pret": 2.32, "capitalizare": 14200, "revenue_growth": 25.0,
        "eps_growth": 30.0, "pe": 4.5, "pe_sector": 8.5, "ev_ebitda": 3.0,
        "fcf_yield": 15.0, "roe": 25.0, "debt_equity": 0.1, "div_yield": 10.0,
        "payout": 90, "scor": 9,
    },
    {
        "simbol": "BRD", "denumire": "BRD Groupe SG", "sector": "Financiar",
        "pret": 18.60, "capitalizare": 12900, "revenue_growth": 10.0,
        "eps_growth": 12.5, "pe": 7.5, "pe_sector": 8.5, "ev_ebitda": 5.8,
        "fcf_yield": 7.0, "roe": 16.0, "debt_equity": 0.8, "div_yield": 6.0,
        "payout": 45, "scor": 7,
    },
    {
        "simbol": "DIGI", "denumire": "Digi Communications", "sector": "Comunicații",
        "pret": 48.50, "capitalizare": 4800, "revenue_growth": 22.0,
        "eps_growth": 35.0, "pe": 12.0, "pe_sector": 15.0, "ev_ebitda": 6.5,
        "fcf_yield": 5.5, "roe": 20.0, "debt_equity": 1.8, "div_yield": 2.0,
        "payout": 25, "scor": 7,
    },
]

# ============================================================================
# CONFIGURARE FOI EXCEL
# ============================================================================
SHEET_ORDER = [
    "Dashboard",
    "Poziții",
    "Tranzacții",
    "Analiză Tehnică",
    "Analiză Fundamentală",
    "Watchlist",
    "Risk Management",
    "Configurare",
]

TAB_COLORS = {
    "Dashboard": "2E75B6",          # Albastru accent
    "Poziții": "27AE60",            # Verde
    "Tranzacții": "E67E22",         # Portocaliu
    "Analiză Tehnică": "8E44AD",    # Violet
    "Analiză Fundamentală": "1ABC9C",  # Teal
    "Watchlist": "F1C40F",          # Auriu
    "Risk Management": "E74C3C",    # Roșu
    "Configurare": "7F8C8D",        # Gri
}

# ============================================================================
# COMISIOANE BROKERI
# ============================================================================
COMISIOANE = {
    "BVB": {"tip": "procent", "valoare": 0.005, "minim": 5.0, "moneda": "RON"},
    "NYSE": {"tip": "per_actiune", "valoare": 0.01, "minim": 1.0, "moneda": "USD"},
    "NASDAQ": {"tip": "per_actiune", "valoare": 0.01, "minim": 1.0, "moneda": "USD"},
    "XETRA": {"tip": "procent", "valoare": 0.003, "minim": 5.0, "moneda": "EUR"},
    "LSE": {"tip": "procent", "valoare": 0.004, "minim": 8.0, "moneda": "GBP"},
    "Euronext": {"tip": "procent", "valoare": 0.003, "minim": 5.0, "moneda": "EUR"},
}
