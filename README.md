# GTO Overeenkomsten Generator API

Backend service voor het GTO Overeenkomsten Platform. Genereert Word + PDF
overeenkomsten op basis van de originele GTO templates en levert deze als ZIP.

## Deploy op Render

1. Maak een nieuw "Web Service" aan op render.com
2. Verbind deze GitHub repository
3. Render herkent automatisch de Dockerfile
4. Environment: Docker
5. Plan: Free
6. Deploy

Na deploy krijg je een URL zoals `https://gto-overeenkomsten-api.onrender.com`

## Endpoints

- `GET /health` - health check
- `POST /generate` - genereert overeenkomsten ZIP

## Let op: gratis Render-tier

De gratis tier "slaapt" na 15 minuten inactiviteit. De eerste request na
inactiviteit duurt daardoor ~30-50 seconden extra (opstarttijd). Daarna
reageert de service weer normaal snel.
