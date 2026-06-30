"""
GTO Overeenkomsten Generator API
==================================
Flask backend die op Render draait. Het React-platform (op Vercel) roept
deze API aan om Word/PDF overeenkomsten te genereren.

Endpoints:
  POST /generate  -> genereert ZIP met Word+PDF overeenkomsten, retourneert als download
  GET  /health     -> health check
"""

import os
import uuid
import tempfile
from pathlib import Path
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS

import generator

app = Flask(__name__)
CORS(app)  # Sta requests toe vanaf het Vercel-platform

OUTPUT_DIR = Path(tempfile.gettempdir()) / "gto_output"
OUTPUT_DIR.mkdir(exist_ok=True)


@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok"})


@app.route("/debug-signatures", methods=["GET"])
def debug_signatures():
    """
    Tijdelijk diagnose-endpoint om te checken of de handtekening-bestanden
    daadwerkelijk aanwezig zijn op de server zoals gedeployed, en of de
    sdt-targeting in de templates klopt. Kan later weer verwijderd worden.
    """
    import os
    info = {
        "signatures_dir": str(generator.SIGNATURES_DIR),
        "signatures_dir_exists": generator.SIGNATURES_DIR.exists(),
        "signatures_dir_contents": [],
        "templates_dir": str(generator.TEMPLATES_DIR),
        "templates_dir_exists": generator.TEMPLATES_DIR.exists(),
        "templates_dir_contents": [],
    }
    if generator.SIGNATURES_DIR.exists():
        info["signatures_dir_contents"] = os.listdir(generator.SIGNATURES_DIR)
    if generator.TEMPLATES_DIR.exists():
        info["templates_dir_contents"] = os.listdir(generator.TEMPLATES_DIR)

    # Test ook de target-detectie op het echte NL ZZP template
    try:
        template_path = generator.TEMPLATES_DIR / generator.TEMPLATE_FILES["nl_zzp"]
        work_unpack = generator.WORK_DIR / "debug_endpoint_test"
        generator.unpack_docx(template_path, work_unpack)
        targets = generator.find_signature_image_targets(work_unpack)
        info["nl_zzp_signature_targets"] = list(targets)
        info["nl_zzp_media_files"] = os.listdir(work_unpack / "word" / "media")
    except Exception as e:
        info["nl_zzp_test_error"] = str(e)

    return jsonify(info)


@app.route("/validate", methods=["POST"])
def validate():
    """
    Controleert of alle benodigde velden aanwezig zijn voordat er gegenereerd
    wordt. Het platform roept dit aan zodra de gebruiker bij stap 4 (tarieven
    & datum) komt, zodat eventuele ontbrekende info (tekenbevoegde, projectnaam,
    etc.) meteen gevraagd kan worden in plaats van pas na het genereren te
    ontdekken dat er lege velden in de overeenkomst staan.

    Verwacht dezelfde payload-structuur als /generate (zonder de datum/tarief-velden,
    die zijn op dit punt nog niet relevant voor de validatie).

    Retourneert: {"compleet": bool, "ontbrekend": [{"veld", "label", "context", "monteur_id"?}]}
    """
    try:
        data = request.get_json(force=True)
        missing = generator.validate_required_fields(
            entiteit=data.get("entiteit", "nl"),
            monteurs=data.get("monteurs", []),
            klant=data.get("klant", {}),
            project=data.get("project", {}),
            tekenbevoegde=data.get("tekenbevoegde", ""),
            opdrachtomschrijving=data.get("opdrachtomschrijving", ""),
        )
        return jsonify({"compleet": len(missing) == 0, "ontbrekend": missing})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/generate", methods=["POST"])
def generate():
    """
    Verwacht JSON body:
    {
      "entiteit": "nl" | "west",
      "monteurs": [{"id": 1, "naam": "...", "handelsnaam": "...", "kvk": "...", "adres": "...", "persnr": "..."}],
      "klant": {"naam": "...", "adres": "...", "kvk": "...", "btw": "...", "contact": "...", "fmail": "..."},
      "project": {"nr": "...", "naam": "...", "werkadres": "..."},
      "tekenbevoegde": "...",
      "startdatum_nl": "DD-MM-JJJJ",
      "startdatum_iso": "JJJJ-MM-DD",
      "handtekendatum_nl": "DD-MM-JJJJ",
      "tarieven_per_monteur": {"1": {"zzp": "45,00", "klant": "54,00", "reisuur": ""}},
      "opdrachtomschrijving": "..."
    }
    """
    try:
        data = request.get_json(force=True)

        required = ["entiteit", "monteurs", "klant", "project", "tekenbevoegde",
                    "startdatum_nl", "startdatum_iso", "handtekendatum_nl",
                    "tarieven_per_monteur", "opdrachtomschrijving"]
        missing = [f for f in required if f not in data]
        if missing:
            return jsonify({"error": f"Ontbrekende velden: {', '.join(missing)}"}), 400

        if data["entiteit"] not in ("nl", "west"):
            return jsonify({"error": "entiteit moet 'nl' of 'west' zijn"}), 400

        # Server-side validatie als laatste vangnet (het platform vraagt dit normaal
        # al af via /validate vóórdat de gebruiker bij de downloadknop komt, maar we
        # checken het hier nogmaals zodat er nooit een overeenkomst met lege
        # verplichte velden wordt gegenereerd, ook niet bij een platform-bugje).
        veld_missing = generator.validate_required_fields(
            entiteit=data["entiteit"], monteurs=data["monteurs"], klant=data["klant"],
            project=data["project"], tekenbevoegde=data["tekenbevoegde"],
            opdrachtomschrijving=data["opdrachtomschrijving"],
        )
        if veld_missing:
            labels = ", ".join(f"{m['label']} ({m['context']})" for m in veld_missing)
            return jsonify({"error": f"Niet alle verplichte velden zijn ingevuld: {labels}", "ontbrekend": veld_missing}), 400

        # tarieven_per_monteur komt als JSON met string keys, converteer naar int waar nodig
        tarieven_raw = data["tarieven_per_monteur"]
        tarieven = {}
        for monteur in data["monteurs"]:
            mid = monteur["id"]
            tarieven[mid] = tarieven_raw.get(str(mid), tarieven_raw.get(mid, {}))

        job_id = str(uuid.uuid4())[:8]
        zip_filename = f"GTO_Overeenkomsten_{job_id}.zip"
        zip_path = OUTPUT_DIR / zip_filename

        generator.generate_full_package(
            entiteit=data["entiteit"],
            monteurs=data["monteurs"],
            klant=data["klant"],
            project=data["project"],
            tekenbevoegde=data["tekenbevoegde"],
            startdatum_nl=data["startdatum_nl"],
            startdatum_iso=data["startdatum_iso"],
            handtekendatum_nl=data["handtekendatum_nl"],
            tarieven_per_monteur=tarieven,
            opdrachtomschrijving=data["opdrachtomschrijving"],
            output_zip_path=str(zip_path),
        )

        download_name = f"Overeenkomsten {data['klant']['naam']} {data['project']['nr']}.zip"

        return send_file(
            zip_path,
            mimetype="application/zip",
            as_attachment=True,
            download_name=download_name,
        )

    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
