"""
advizeo ML Satz Converter – Flask Web Application
==================================================
Locally-hosted web server that:
  1. Accepts BFW DTA ML Satz file uploads
  2. Parses L/M/B records
  3. Serves advizeo-compatible Excel import templates for download

Run: python app.py
Open: http://localhost:5000
"""

import os
import uuid
from io import BytesIO

from flask import (
    Flask,
    jsonify,
    render_template,
    request,
    send_file,
)

from parser import parse_ml_satz
from excel_generator import (
    count_rows,
    generate_building_structure,
    generate_create_devices,
    generate_create_tenants,
    generate_update_devices,
    generate_update_tenants,
)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50 MB

# In-memory session storage  {session_id: parsed_dict}
_sessions: dict[str, dict] = {}

# ---------------------------------------------------------------------------
# Template generators registry
# ---------------------------------------------------------------------------

TEMPLATES = {
    "building_structure": (
        generate_building_structure,
        "create_building_structure_import.xlsx",
        "Gebäudestruktur / Building Structure",
    ),
    "create_tenants": (
        generate_create_tenants,
        "create_tenants_import.xlsx",
        "Mieter anlegen / Create Tenants",
    ),
    "update_tenants": (
        generate_update_tenants,
        "update_tenants_import.xlsx",
        "Mieter aktualisieren / Update Tenants",
    ),
    "create_devices": (
        generate_create_devices,
        "create_devices_import.xlsx",
        "Geräte anlegen / Create Devices",
    ),
    "update_devices": (
        generate_update_devices,
        "update_devices_import.xlsx",
        "Geräte aktualisieren / Update Devices",
    ),
}


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/upload", methods=["POST"])
def upload():
    if "file" not in request.files:
        return jsonify({"error": "Keine Datei übergeben / No file provided"}), 400

    file = request.files["file"]
    if not file or file.filename == "":
        return jsonify({"error": "Keine Datei ausgewählt / No file selected"}), 400

    # Read raw bytes and decode with fallback encodings
    raw_bytes = file.read()
    text = None
    for enc in ("utf-8", "latin-1", "cp1252", "iso-8859-1"):
        try:
            text = raw_bytes.decode(enc)
            break
        except (UnicodeDecodeError, LookupError):
            continue

    if text is None:
        return jsonify(
            {"error": "Zeichenkodierung nicht erkannt / Could not decode file encoding"}
        ), 400

    try:
        parsed = parse_ml_satz(text)
    except Exception as exc:
        return jsonify({"error": f"Parsing-Fehler: {exc}"}), 500

    session_id = str(uuid.uuid4())
    _sessions[session_id] = parsed

    row_counts = count_rows(parsed)
    preview    = _build_preview(parsed)

    return jsonify(
        {
            "session_id": session_id,
            "filename": file.filename,
            "records": {
                "a_saetze": len(parsed.get("a_saetze", [])),
                "l_saetze": len(parsed.get("l_saetze", [])),
                "m_saetze": len(parsed.get("m_saetze", [])),
                "b_saetze": len(parsed.get("b_saetze", [])),
                "d_saetze": len(parsed.get("d_saetze", [])),
            },
            "row_counts": row_counts,
            "errors": parsed.get("errors", []),
            "preview": preview,
        }
    )


@app.route("/api/download/<session_id>/<template_type>")
def download(session_id: str, template_type: str):
    if session_id not in _sessions:
        return (
            jsonify(
                {
                    "error": (
                        "Sitzung nicht gefunden. Bitte laden Sie die Datei erneut hoch. "
                        "/ Session not found – please re-upload the file."
                    )
                }
            ),
            404,
        )

    if template_type not in TEMPLATES:
        return jsonify({"error": f"Unbekannter Template-Typ: {template_type}"}), 400

    parsed = _sessions[session_id]
    gen_fn, filename, _ = TEMPLATES[template_type]

    try:
        excel_bytes = gen_fn(parsed)
    except Exception as exc:
        return jsonify({"error": f"Generierungsfehler: {exc}"}), 500

    return send_file(
        BytesIO(excel_bytes),
        mimetype=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
        as_attachment=True,
        download_name=filename,
    )


@app.route("/api/sample")
def sample():
    """Return a sample ML Satz file for demonstration."""
    sample_path = os.path.join(app.static_folder, "sample.ml")
    if not os.path.exists(sample_path):
        return jsonify({"error": "Beispieldatei nicht gefunden"}), 404
    return send_file(sample_path, as_attachment=True, download_name="sample.ml")


# ---------------------------------------------------------------------------
# Preview builder
# ---------------------------------------------------------------------------

def _build_preview(parsed: dict) -> dict:
    properties = []
    for l in parsed.get("l_saetze", []):
        properties.append(
            {
                "property_number": l.property_number,
                "external_id":     l.external_id,
                "name":            l.real_estate_name,
                "street":          l.strasse,
                "postal_code":     l.postleitzahl,
                "city":            l.stadt,
                "country":         l.laendercode or "de",
            }
        )

    # Build property_number → external_id lookup for tenant rows
    prop_ext_lookup = {l.property_number: l.external_id for l in parsed.get("l_saetze", [])}

    tenants = []
    for m in parsed.get("m_saetze", []):
        tenants.append(
            {
                "property_number":    m.property_number,
                "real_estate_ext_id": prop_ext_lookup.get(m.property_number) or m.property_number,
                "unit_external_id":   m.estate_unit_external_id,
                "apartment_number":   m.apartment_number,
                "name":               "Leerstand" if m.is_vacant else m.tenant_name,
                "move_in":            m.einzugsdatum,
                "move_out":           m.auszugsdatum,
                "vacant":             m.is_vacant,
            }
        )

    devices = []
    for b in parsed.get("b_saetze", []):
        devices.append(
            {
                "property_number": b.property_number,
                "device_id":       b.geraete_nr,
                "device_type":     b.device_type_advizeo,
                "brennstoffart":   b.brennstoffart,
                "start_date":      b.abrechnungszeitraum_start,
                "end_date":        b.abrechnungszeitraum_ende,
            }
        )

    return {"properties": properties, "tenants": tenants, "devices": devices}


# ---------------------------------------------------------------------------
# Entrypoint
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    print("\n" + "=" * 60)
    print("  advizeo ML Satz Converter")
    print("  http://localhost:5000")
    print("=" * 60 + "\n")
    app.run(debug=False, host="0.0.0.0", port=5000)
