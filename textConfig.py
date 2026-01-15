TEXTS = {
    "de": {
        "filename": "meine-finanzen",
        "header_title": "MEINE FINANZEN",
        "assets_title": "AKTUELLES VERMÖGEN (HEUTE)",
        "assets_labels": ["Girokonto", "Tagesgeld", "Bargeld", "Depot", "Netto-Einkommen"],
        
        # KPIs
        "kpi_fix_real": "FIXKOSTEN REAL (mtl.)",
        "kpi_buffer": "RÜCKLAGEN PUFFER (mtl.)",
        "kpi_runway": "NOTGROSCHEN-DAUER",
        "kpi_save_rate": "SPARQUOTE",
        "kpi_total": "AKTUELLES VERMÖGEN",
        "unit_months": "Monate",
        
        # Vorschau
        "preview_title": "KONTOSTAND-VORSCHAU",
        "months": ["Jan","Feb","Mrz","Apr","Mai","Jun","Jul","Aug","Sep","Okt","Nov","Dez"],
        
        # Tabellen Header
        "col_item": "Posten",
        "col_freq": "Turnus",
        "col_amount": "Betrag",
        "col_due": "Fällig",
        "col_exact": "Exakt (mtl.)",
        "col_buffer": "Rücklage (+Puffer)",
        "col_goal": "Ziel",
        "col_class": "Anlage",
        
        # Kategorien Titel
        "cat_living": "🏠 WOHNEN & LEBEN",
        "cat_digital": "💻 DIGITALES & ABOS",
        "cat_insurance": "🛡️ VERSICHERUNG",
        "cat_invest": "📈 SPARPLÄNE",
        "cat_log": "📊 VERMÖGENS-VERLAUF",
        
        # Log Spalten
        "log_cols": ["Monat", "Giro", "Tagesgeld", "Bargeld", "Depot", "Investiert", "Unterschied"],
        
        # Dashboard & Info Center
        "dash_title": "MEINE FINANZEN - DASHBOARD",
        "btn_open": "ÖFFNEN",
        "chart_title": "Gesamtvermögens-Entwicklung",
        "chart_y": "Vermögen (€)",
        "chart_x": "Zeitraum",
        "info_title": "📘 HANDBUCH & GLOSSAR",
        
        # Manual Content (Titel, Text)
        "manual": [
            ("🚀 SCHNELLSTART", "Gehe ins Jahres-Sheet. Ganz rechts 'VERMÖGENS-VERLAUF'. Trage bei 'Jan' Startwerte ein."),
            ("🔄 AUTO-AKTUALISIERUNG", "Die Box oben links zeigt automatisch die Werte des aktuellen Monats aus der Verlaufstabelle."),
            ("📅 DEINE TO-DO", "Einmal im Monat im Verlauf (rechts) neue Stände eintragen."),
            ("💡 INTERVALL-ZAHLUNGEN", "Turnus > 1? Dann bei 'Fällig' Monate auflisten: 'Feb, Mai, Aug, Nov'."),
            ("📉 NOTGROSCHEN-DAUER", "Reichweite: (Bargeld + Tagesgeld) / reale Fixkosten."),
            ("💰 SPARQUOTE", "(Sparpläne + Überschuss) / Netto-Einkommen."),
            ("🛡️ PUFFER-RÜCKLAGE", "Fixkosten durch 12 + 5% Puffer.")
        ],
        
        # Beispiel Daten
        "sample_rent": "Miete",
        "sample_gez": "Rundfunkbeitrag",
        "sample_netflix": "Streaming",
        "sample_hosting": "Server/Cloud",
        "sample_kfz": "KFZ-Steuer",
        "sample_etf": "Welt-ETF"
    },
    
    "en": {
        "filename": "my-finances",
        "header_title": "MY FINANCES",
        "assets_title": "CURRENT ASSETS (TODAY)",
        "assets_labels": ["Checking", "Savings", "Cash", "Portfolio", "Net Income"],
        
        "kpi_fix_real": "REAL FIXED COSTS (mth)",
        "kpi_buffer": "BUFFER SAVINGS (mth)",
        "kpi_runway": "RUNWAY (CASH)",
        "kpi_save_rate": "SAVINGS RATE",
        "kpi_total": "TOTAL NET WORTH",
        "unit_months": "Months",
        
        "preview_title": "CASHFLOW PREVIEW",
        "months": ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"],
        
        "col_item": "Item",
        "col_freq": "Freq",
        "col_amount": "Amount",
        "col_due": "Due",
        "col_exact": "Exact (mth)",
        "col_buffer": "Buffer (+5%)",
        "col_goal": "Goal",
        "col_class": "Asset",
        
        "cat_living": "🏠 LIVING & LIFE",
        "cat_digital": "💻 DIGITAL & SUBSCRIPTIONS",
        "cat_insurance": "🛡️ INSURANCE",
        "cat_invest": "📈 INVESTMENTS",
        "cat_log": "📊 WEALTH LOG",
        
        "log_cols": ["Month", "Checking", "Savings", "Cash", "Portfolio", "Invested", "Diff"],
        
        "dash_title": "MY FINANCES - DASHBOARD",
        "btn_open": "OPEN",
        "chart_title": "Total Net Worth History",
        "chart_y": "Net Worth (€)",
        "chart_x": "Timeframe",
        "info_title": "📘 MANUAL & GLOSSARY",
        
        "manual": [
            ("🚀 QUICK START", "Go to the Year-Sheet. Far right table 'WEALTH LOG'. Enter starting values for 'Jan'."),
            ("🔄 AUTO-UPDATE", "The box top-left automatically pulls current month's values from the log table."),
            ("📅 YOUR TO-DO", "Update your account balances in the log table once a month."),
            ("💡 INTERVAL PAYMENTS", "Freq > 1? List months in 'Due': 'Feb, May, Aug, Nov'."),
            ("📉 RUNWAY", "Survival time: (Cash + Savings) / Real fixed costs."),
            ("💰 SAVINGS RATE", "(Investments + Surplus) / Net Income."),
            ("🛡️ BUFFER", "Real fixed costs + 5% safety margin.")
        ],
        
        "sample_rent": "Rent",
        "sample_gez": "TV License",
        "sample_netflix": "Streaming",
        "sample_hosting": "Hosting",
        "sample_kfz": "Car Tax",
        "sample_etf": "Global ETF"
    }
}