# Gunicorn configuratie voor GTO Overeenkomsten Backend
# Verhoogde timeout omdat LibreOffice PDF-conversie tijd nodig heeft

timeout = 300        # 5 minuten — genoeg voor meerdere PDF-conversies
workers = 1          # 1 worker (gratis tier heeft beperkt RAM)
bind = "0.0.0.0:5000"
