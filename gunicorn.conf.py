# Gunicorn configuratie voor GTO Overeenkomsten Backend
# Verhoogde timeout omdat LibreOffice PDF-conversie tijd nodig heeft

timeout = 600        # 10 minuten — voor grote batches (12+ monteurs × 2 docs + PDF)
workers = 1          # 1 worker (gratis tier heeft beperkt RAM)
bind = "0.0.0.0:5000"
