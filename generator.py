"""
GTO Overeenkomst Generator - Definitieve versie
Bevat: SDT-vulling, handtekening-injectie, artikel-injectie, REGON-detectie
"""

import re
import os
import shutil
import zipfile
import subprocess
import tempfile as _tempfile
from pathlib import Path

# ── Paden ──────────────────────────────────────────────────────────────────────
TEMPLATES_DIR = Path(__file__).parent / "templates"
WORK_DIR = Path(_tempfile.gettempdir()) / "gto_work"
WORK_DIR.mkdir(parents=True, exist_ok=True)
SIGNATURES_DIR = Path(__file__).parent / "signatures"
SIGNATURE_FILES = {
    "nl": "handtekening_nl.png",
    "west": "handtekening_west.png",
}

# ── SDT mappings ───────────────────────────────────────────────────────────────
SDT_MAP_NL_ZZP = {
    0: "projectnr_combo", 2: "monteur_handelsnaam", 3: "monteur_adres",
    4: "monteur_naam", 5: "klant_naam", 6: "project_naam",
    7: "project_werkadres", 8: "project_nr", 10: "monteur_naam2",
    11: "monteur_handelsnaam2", 12: "monteur_kvk", 13: "opdrachtomschrijving",
    14: "startdatum", 16: "uurtarief_zzp", 18: "handtekendatum",
    20: "monteur_handelsnaam3",
}
SDT_MAP_WEST_ZZP = {
    0: "projectnr_combo", 2: "monteur_handelsnaam", 3: "monteur_adres",
    4: "monteur_naam", 5: "klant_naam", 6: "project_naam",
    7: "project_werkadres", 8: "project_nr", 10: "monteur_naam2",
    11: "monteur_handelsnaam2", 12: "monteur_kvk", 13: "opdrachtomschrijving",
    14: "startdatum", 16: "uurtarief_zzp", 18: "handtekendatum",
    20: "monteur_handelsnaam3",
}
SDT_MAP_NL_KLANT = {
    0: "projectnr_combo", 1: "klant_naam", 2: "klant_adres",
    3: "tekenbevoegde", 5: "klant_naam", 6: "klant_adres",
    7: "klant_kvk", 8: "klant_btw", 9: "klant_contact",
    10: "klant_werkadres", 11: "klant_fmail",
    14: "monteur_naam", 15: "monteur_handelsnaam", 16: "monteur_kvk",
    17: "opdrachtomschrijving", 18: "project_nr", 19: "startdatum",
    20: "einddatum", 21: "uurtarief_klant",
    24: "tekenbevoegde", 26: "handtekendatum",
}
SDT_MAP_WEST_KLANT = {
    0: "projectnr_combo", 1: "klant_naam", 2: "klant_adres",
    3: "tekenbevoegde", 5: "klant_naam", 6: "klant_adres",
    7: "klant_kvk", 8: "klant_btw", 9: "klant_contact",
    10: "klant_werkadres", 11: "klant_fmail",
    13: "monteur_naam", 14: "monteur_handelsnaam", 15: "monteur_kvk",
    16: "opdrachtomschrijving", 17: "project_nr", 18: "startdatum",
    19: "einddatum", 20: "uurtarief_klant",
    22: "tekenbevoegde", 24: "handtekendatum",
}
TEMPLATE_FILES = {
    "nl_zzp":     "Vernieuwde_overeenkomst_ZZP_2026_NEDERLAND.docx",
    "west_zzp":   "Vernieuwde_overeenkomst_ZZP_2026_WEST.docx",
    "nl_klant":   "NIEUW_projectovereenkomst_KLANT_2026_Nederland.docx",
    "west_klant": "NIEUW_projectovereenkomst_KLANT_2026_West.docx",
}
SDT_MAPS = {
    "nl_zzp": SDT_MAP_NL_ZZP, "west_zzp": SDT_MAP_WEST_ZZP,
    "nl_klant": SDT_MAP_NL_KLANT, "west_klant": SDT_MAP_WEST_KLANT,
}
GROTE_FONT_VELDEN = {"monteur_handelsnaam3"}


# ── Docx unpack / pack ─────────────────────────────────────────────────────────
def unpack_docx(docx_path, work_dir):
    work_dir = Path(work_dir)
    if work_dir.exists():
        shutil.rmtree(work_dir)
    work_dir.mkdir(parents=True)
    with zipfile.ZipFile(docx_path, "r") as zf:
        zf.extractall(work_dir)
    return work_dir


def pack_docx(work_dir, output_path):
    work_dir = Path(work_dir)
    output_path = Path(output_path)
    if output_path.exists():
        output_path.unlink()
    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zf:
        all_files = sorted(
            [f for f in work_dir.rglob("*") if f.is_file()],
            key=lambda f: (0 if f.relative_to(work_dir).as_posix() == "[Content_Types].xml" else 1,
                           f.relative_to(work_dir).as_posix())
        )
        for f in all_files:
            zf.write(f, f.relative_to(work_dir).as_posix())
    return output_path


# ── XML helpers ────────────────────────────────────────────────────────────────
def escape_xml(text):
    if text is None: text = ""
    text = str(text)
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def fixed_rpr(size_half_points=14):
    return (f'<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>'
            f'<w:sz w:val="{size_half_points}"/><w:szCs w:val="{size_half_points}"/></w:rPr>')


def find_balanced_sdt_blocks(xml_content):
    tag_pattern = re.compile(r"<(/?)w:sdt(\s|>|/>)")
    depth = 0
    top_level_blocks = []
    current_start = None
    for m in tag_pattern.finditer(xml_content):
        is_closing = m.group(1) == "/"
        if not is_closing:
            if depth == 0:
                current_start = m.start()
            depth += 1
        else:
            depth -= 1
            if depth == 0 and current_start is not None:
                close_idx = xml_content.find(">", m.start())
                top_level_blocks.append((current_start, close_idx + 1))
                current_start = None
    return top_level_blocks


def fill_sdt_content(xml_content, sdt_index, new_text, blocks_cache=None, font_size=14):
    blocks = blocks_cache if blocks_cache is not None else find_balanced_sdt_blocks(xml_content)
    if sdt_index >= len(blocks):
        return xml_content
    start, end = blocks[sdt_index]
    sdt_block = xml_content[start:end]
    escaped_text = escape_xml(new_text)
    needs_preserve = (new_text == "") or (new_text != new_text.strip())
    space_attr = ' xml:space="preserve"' if needs_preserve else ""
    rpr_xml = fixed_rpr(font_size)
    content_tag_pattern = re.compile(r"<(/?)w:sdtContent(\s|>|/>)")
    depth = 0
    content_start = None
    content_end = None
    for m in content_tag_pattern.finditer(sdt_block):
        is_closing = m.group(1) == "/"
        if not is_closing:
            if depth == 0:
                content_start = sdt_block.find(">", m.start()) + 1
            depth += 1
        else:
            depth -= 1
            if depth == 0:
                content_end = m.start()
                break
    if content_start is None or content_end is None:
        return xml_content
    inner = sdt_block[content_start:content_end]
    has_block_element = bool(re.search(r"<w:(tc|p)\b", inner))
    new_run = f"<w:r>{rpr_xml}<w:t{space_attr}>{escaped_text}</w:t></w:r>"
    if has_block_element:
        run_pattern = re.compile(r"<w:r\b[^>]*>(?:(?!</w:r>).)*?</w:r>", re.DOTALL)
        runs = list(run_pattern.finditer(inner))
        if not runs:
            return xml_content
        pieces = []
        last_end = 0
        for i, r in enumerate(runs):
            pieces.append(inner[last_end:r.start()])
            if i == 0:
                pieces.append(new_run)
            last_end = r.end()
        pieces.append(inner[last_end:])
        new_inner = "".join(pieces)
    else:
        new_inner = new_run
    new_sdt_block = sdt_block[:content_start] + new_inner + sdt_block[content_end:]
    return xml_content[:start] + new_sdt_block + xml_content[end:]


# ── Handtekening ───────────────────────────────────────────────────────────────
def find_signature_image_targets(work_unpack_dir):
    doc_xml_path = work_unpack_dir / "word" / "document.xml"
    rels_path = work_unpack_dir / "word" / "_rels" / "document.xml.rels"
    with open(doc_xml_path, "r", encoding="utf-8") as f:
        xml_content = f.read()
    with open(rels_path, "r", encoding="utf-8") as f:
        rels_content = f.read()
    anchor_pattern = re.compile(r"<wp:anchor.*?</wp:anchor>", re.DOTALL)
    embed_ids = set()
    for m in anchor_pattern.finditer(xml_content):
        for rid_match in re.finditer(r'r:embed="(rId\d+)"', m.group()):
            embed_ids.add(rid_match.group(1))
    targets = set()
    for rid in embed_ids:
        m = re.search(rf'<Relationship Id="{rid}"[^>]*Target="media/([^"]+)"', rels_content)
        if m:
            targets.add(m.group(1))
    return targets


def inject_signature_image(work_unpack_dir, entiteit):
    sig_filename = SIGNATURE_FILES.get(entiteit)
    if not sig_filename:
        return
    sig_path = SIGNATURES_DIR / sig_filename
    if not sig_path.exists():
        return
    targets = find_signature_image_targets(work_unpack_dir)
    media_dir = work_unpack_dir / "word" / "media"
    for target_filename in targets:
        target_path = media_dir / target_filename
        if target_path.exists():
            shutil.copy(sig_path, target_path)


# ── Artikel-injectie ───────────────────────────────────────────────────────────
def artikel_to_xml(artikel):
    """
    Converteer een artikel-dict naar Word XML paragrafen met correcte nummering.
    Opmaak:
    - Artikel-header: vetgedrukt Arial 9pt, bijv. '1.  De opdracht'
    - Sub-artikelen: Arial 8pt, ingesprongen, genummerd als '1.1  tekst...'
    - Opsommingspunten (streepje of letter a./b./c.): iets dieper ingesprongen
    """
    import re as _re
    art_nr = artikel['nr']
    xml_parts = []

    # Artikel-header (vetgedrukt)
    titel = escape_xml(f"{art_nr}.  {artikel['titel']}")
    xml_parts.append(
        '<w:p><w:pPr><w:spacing w:before="160" w:after="60"/></w:pPr>'
        '<w:r><w:rPr><w:b/><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>'
        '<w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr>'
        f'<w:t>{titel}</w:t></w:r></w:p>'
    )

    for i, sub in enumerate(artikel.get('subaartikelen', []), 1):
        raw_tekst = sub['tekst']

        # Detecteer of dit een opsommingspunt is (begint met streepje, letter+punt, of 'a.'/'b.')
        is_opsomming = bool(
            _re.match(r'^[-–—]\s', raw_tekst) or   # streepje
            _re.match(r'^[a-z]\.\s', raw_tekst) or            # a. b. c.
            _re.match(r'^[A-Z]\.\s', raw_tekst)               # A. B. C.
        )

        if is_opsomming:
            # Opsommingspunt: dieper ingesprongen, geen extra nummering
            tekst = escape_xml(raw_tekst)
            xml_parts.append(
                '<w:p><w:pPr><w:ind w:left="720" w:hanging="360"/>'
                '<w:spacing w:before="40" w:after="40"/></w:pPr>'
                '<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>'
                '<w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr>'
                f'<w:t xml:space="preserve">{tekst}</w:t></w:r></w:p>'
            )
        else:
            # Normaal sub-artikel met nummering: art_nr.i
            nummer = escape_xml(f"{art_nr}.{i}")
            tekst = escape_xml(raw_tekst)
            xml_parts.append(
                '<w:p><w:pPr><w:ind w:left="540" w:hanging="360"/>'
                '<w:spacing w:before="60" w:after="60"/></w:pPr>'
                '<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>'
                '<w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr>'
                f'<w:t xml:space="preserve">{nummer}  </w:t></w:r>'
                '<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>'
                '<w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr>'
                f'<w:t xml:space="preserve">{tekst}</w:t></w:r></w:p>'
            )

    return ''.join(xml_parts)


def inject_artikelen(xml_content, artikelen, template_key):
    """
    Vervang de vaste artikel-paragrafen door de opgegeven artikelen.
    Gebruikt de EERSTE artikel-titel als anker om de beginpositie te vinden,
    en de LAATSTE tabel als anker voor de eindpositie (handtekentabel).
    Dit is robuust omdat het niet afhangt van HANDTEKENING-labels die soms
    in datatabellen staan en daardoor op de verkeerde plek kunnen matchen.
    """
    if not artikelen:
        return xml_content

    nieuwe_artikel_xml = ''.join(artikel_to_xml(a) for a in artikelen)
    eerste_titel = artikelen[0].get('titel', '') if artikelen else ''

    # Vind de EERSTE artikel-header paragraaf in het document
    para_pattern = re.compile(r'<w:p\b[^>]*>(?:(?!</w:p>).)*?</w:p>', re.DOTALL)
    paragraphs = list(para_pattern.finditer(xml_content))

    artikel_begin_pos = None

    # Probeer eerst de exacte titel van het eerste artikel te vinden
    for m in paragraphs:
        texts = re.findall(r'<w:t[^>]*>([^<]*)</w:t>', m.group())
        full = ''.join(texts).strip()
        if eerste_titel and full == eerste_titel:
            artikel_begin_pos = m.start()
            break

    # Fallback voor ZZP: zoek puntjes-lijn en begin ACHTER die lijn
    if artikel_begin_pos is None:
        for m in paragraphs:
            texts = re.findall(r'<w:t[^>]*>([^<]*)</w:t>', m.group())
            full = ''.join(texts).strip()
            if '\u2026' in full or '……' in full or '......' in full:
                artikel_begin_pos = m.end()
                break

    if artikel_begin_pos is None:
        return xml_content  # Veilig: ongewijzigd laten

    # Vind de LAATSTE tabel in het document (altijd de handtekentabel)
    tbl_positions = [m.start() for m in re.finditer(r'<w:tbl>', xml_content)]
    if not tbl_positions:
        return xml_content
    handteken_tbl_pos = tbl_positions[-1]

    if artikel_begin_pos >= handteken_tbl_pos:
        return xml_content  # Sanity check

    return (
        xml_content[:artikel_begin_pos] +
        '\n' + nieuwe_artikel_xml + '\n' +
        xml_content[handteken_tbl_pos:]
    )


# ── Generator ──────────────────────────────────────────────────────────────────
def generate_docx(template_key, data, output_path, artikelen=None):
    """Genereer een ingevuld .docx bestand. artikelen optioneel."""
    template_path = TEMPLATES_DIR / TEMPLATE_FILES[template_key]
    sdt_map = SDT_MAPS[template_key]
    work_unpack = WORK_DIR / f"unpack_{template_key}_{os.getpid()}_{id(data)}"
    unpack_docx(template_path, work_unpack)
    entiteit_voor_handtekening = "west" if template_key.startswith("west") else "nl"
    inject_signature_image(work_unpack, entiteit_voor_handtekening)
    doc_xml_path = work_unpack / "word" / "document.xml"
    with open(doc_xml_path, "r", encoding="utf-8") as f:
        xml_content = f.read()
    for idx in sorted(sdt_map.keys(), reverse=True):
        field_name = sdt_map[idx]
        value = data.get(field_name, "")
        font_size = 22 if field_name in GROTE_FONT_VELDEN else 14
        blocks = find_balanced_sdt_blocks(xml_content)
        xml_content = fill_sdt_content(xml_content, idx, value, blocks_cache=blocks, font_size=font_size)
    if artikelen:
        xml_content = inject_artikelen(xml_content, artikelen, template_key)
    with open(doc_xml_path, "w", encoding="utf-8") as f:
        f.write(xml_content)
    for rel_path, tag_pat in [
        (work_unpack / "word" / "_rels" / "settings.xml.rels", r"<Relationship[^>]*attachedTemplate[^>]*/>"),
        (work_unpack / "word" / "settings.xml", r"<w:attachedTemplate[^/]*/>"),
    ]:
        if rel_path.exists():
            with open(rel_path, "r", encoding="utf-8") as f:
                c = f.read()
            if "attachedTemplate" in c:
                c = re.sub(tag_pat, "", c)
                with open(rel_path, "w", encoding="utf-8") as f:
                    f.write(c)
    pack_docx(work_unpack, output_path)
    shutil.rmtree(work_unpack, ignore_errors=True)
    return output_path


def docx_to_pdf(docx_path, output_dir):
    """Converteer docx naar pdf via LibreOffice. Geeft pad terug naar de pdf."""
    docx_path = Path(docx_path)
    output_dir = Path(output_dir)
    result = subprocess.run(
        ["soffice", "--headless", "--convert-to", "pdf",
         "--outdir", str(output_dir), str(docx_path)],
        capture_output=True, timeout=90
    )
    pdf_path = output_dir / (docx_path.stem + ".pdf")
    if not pdf_path.exists():
        # LibreOffice kan soms de naam aanpassen; zoek het nieuwste PDF-bestand
        pdfs = sorted(output_dir.glob("*.pdf"), key=lambda p: p.stat().st_mtime, reverse=True)
        if pdfs:
            return pdfs[0]
        raise FileNotFoundError(
            f"PDF niet gevonden na conversie van {docx_path.name}. "
            f"LibreOffice stdout: {result.stdout.decode(errors='replace')[:200]}"
        )
    return pdf_path


# ── Hulpfuncties ───────────────────────────────────────────────────────────────
def format_uurtarief(tarief, reisuur_tekst=None, met_euroteken=False):
    if not tarief or not str(tarief).strip():
        return ""
    raw = "".join(c for c in str(tarief).strip() if c != "\u20ac").strip()
    raw = raw.replace(".", ",")
    try:
        numeric = float(raw.replace(",", "."))
        formatted = f"{numeric:.2f}".replace(".", ",")
    except ValueError:
        formatted = raw
    if met_euroteken:
        formatted = "\u20ac " + formatted
    if reisuur_tekst and str(reisuur_tekst).strip():
        return f"{formatted} + {str(reisuur_tekst).strip()}"
    return formatted


def format_kvk_or_regon(kvk):
    if not kvk:
        return kvk
    s = str(kvk).strip()
    if s.upper().startswith("REGON"):
        return s
    digits_only = "".join(c for c in s if c.isdigit())
    if len(digits_only) in (9, 14) and digits_only == s.replace(" ", ""):
        return f"REGON {digits_only}"
    return s


def build_filename(doc_type, entiteit, startdatum_iso, persnr, klantnaam, projectnr):
    klantnaam_clean = re.sub(r'[\\/:*?"<>|]', "", klantnaam).strip()
    if doc_type == "zzp":
        return f"{startdatum_iso} PO {persnr} - {klantnaam_clean} {projectnr}"
    else:
        return f"{startdatum_iso} PO {klantnaam_clean} - {persnr} {projectnr}"


def build_field_data(entiteit, monteur, klant, project, tekenbevoegde,
                     startdatum_nl, handtekendatum_nl, uurtarief_zzp, uurtarief_klant,
                     reisuur_tekst, opdrachtomschrijving):
    persnr = monteur["persnr"]
    common = {
        "projectnr_combo":      f"{project['nr']} + {persnr}",
        "klant_naam":           klant["naam"],
        "klant_adres":          klant.get("adres", ""),
        "klant_kvk":            klant.get("kvk", ""),
        "klant_btw":            klant.get("btw", ""),
        "klant_contact":        klant.get("contact", ""),
        "klant_werkadres":      project.get("werkadres", klant.get("adres", "")),
        "klant_fmail":          klant.get("fmail", ""),
        "tekenbevoegde":        tekenbevoegde,
        "monteur_naam":         monteur["naam"],
        "monteur_handelsnaam":  monteur["handelsnaam"],
        "monteur_adres":        monteur.get("adres", ""),
        "monteur_kvk":          format_kvk_or_regon(monteur.get("kvk", "")),
        "opdrachtomschrijving": opdrachtomschrijving,
        "project_nr":           project["nr"],
        "project_naam":         project.get("naam", ""),
        "project_werkadres":    project.get("werkadres", ""),
        "startdatum":           startdatum_nl,
        "handtekendatum":       handtekendatum_nl,
        "einddatum":            "3 tot 6 maanden",
    }
    zzp_data = {**common,
        "monteur_naam2":        monteur["naam"],
        "monteur_handelsnaam2": monteur["handelsnaam"],
        "monteur_handelsnaam3": monteur["handelsnaam"],
        "uurtarief_zzp":        format_uurtarief(uurtarief_zzp, reisuur_tekst, met_euroteken=False),
    }
    klant_data = {**common,
        "uurtarief_klant":      format_uurtarief(uurtarief_klant, reisuur_tekst, met_euroteken=True),
    }
    return zzp_data, klant_data


def validate_required_fields(entiteit, monteurs, klant, project, tekenbevoegde, opdrachtomschrijving):
    missing = []
    if not tekenbevoegde or not str(tekenbevoegde).strip():
        missing.append({"veld": "tekenbevoegde", "label": "Tekenbevoegde van de klant",
                        "context": klant.get("naam", "")})
    for key, label in [("naam", "Klantnaam"), ("adres", "Klantadres"), ("kvk", "KvK-nummer klant")]:
        if not klant.get(key, "").strip():
            missing.append({"veld": f"klant.{key}", "label": label, "context": klant.get("naam", "")})
    for key, label in [("nr", "Projectnummer"), ("naam", "Projectnaam"), ("werkadres", "Werkadres project")]:
        if not project.get(key, "").strip():
            missing.append({"veld": f"project.{key}", "label": label,
                            "context": project.get("naam", project.get("nr", ""))})
    if not opdrachtomschrijving or not str(opdrachtomschrijving).strip():
        missing.append({"veld": "opdrachtomschrijving", "label": "Opdrachtomschrijving",
                        "context": project.get("naam", "")})
    for m in monteurs:
        mnaam = m.get("naam", "(naam onbekend)")
        for key, label in [("naam", "Naam"), ("handelsnaam", "Handelsnaam"), ("kvk", "KvK-nummer")]:
            if not m.get(key, "").strip():
                missing.append({"veld": f"monteur.{key}", "label": f"{label} monteur",
                                "context": mnaam, "monteur_id": m.get("id")})
        if not m.get("persnr", "").strip():
            entiteit_label = "GTO West" if entiteit == "west" else "GTO Nederland"
            missing.append({"veld": "monteur.persnr", "label": f"Persoonsnummer ({entiteit_label})",
                            "context": mnaam, "monteur_id": m.get("id")})
    return missing


def generate_full_package(entiteit, monteurs, klant, project, tekenbevoegde,
                          startdatum_nl, startdatum_iso, handtekendatum_nl,
                          tarieven_per_monteur, opdrachtomschrijving, output_zip_path,
                          artikelen=None):
    """
    Genereer ZIP met Word+PDF voor alle monteurs.
    artikelen: dict met keys '{entiteit}_zzp' en '{entiteit}_klant',
    elk een lijst van artikel-dicts. Als leeg: originele template-artikelen.
    """
    template_zzp_key = f"{entiteit}_zzp"
    template_klant_key = f"{entiteit}_klant"
    tmp_dir = WORK_DIR / f"package_{os.getpid()}_{id(monteurs)}"
    tmp_dir.mkdir(parents=True, exist_ok=True)
    multi = len(monteurs) > 1

    zzp_artikelen = (artikelen or {}).get(template_zzp_key) or \
                    (artikelen or {}).get("zzp") or []
    kl_artikelen  = (artikelen or {}).get(template_klant_key) or \
                    (artikelen or {}).get("klant") or []

    try:
        with zipfile.ZipFile(output_zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
            for monteur in monteurs:
                tarieven = tarieven_per_monteur.get(monteur["id"], {})
                zzp_data, klant_data = build_field_data(
                    entiteit=entiteit, monteur=monteur, klant=klant, project=project,
                    tekenbevoegde=tekenbevoegde, startdatum_nl=startdatum_nl,
                    handtekendatum_nl=handtekendatum_nl,
                    uurtarief_zzp=tarieven.get("zzp", ""),
                    uurtarief_klant=tarieven.get("klant", ""),
                    reisuur_tekst=tarieven.get("reisuur", ""),
                    opdrachtomschrijving=opdrachtomschrijving,
                )
                persnr = monteur["persnr"]
                zzp_fn = build_filename("zzp", entiteit, startdatum_iso, persnr, klant["naam"], project["nr"])
                kl_fn  = build_filename("klant", entiteit, startdatum_iso, persnr, klant["naam"], project["nr"])
                zzp_docx = tmp_dir / f"{zzp_fn}.docx"
                kl_docx  = tmp_dir / f"{kl_fn}.docx"
                generate_docx(template_zzp_key, zzp_data, zzp_docx,
                              artikelen=zzp_artikelen if zzp_artikelen else None)
                generate_docx(template_klant_key, klant_data, kl_docx,
                              artikelen=kl_artikelen if kl_artikelen else None)
                zzp_pdf = docx_to_pdf(zzp_docx, tmp_dir)
                kl_pdf  = docx_to_pdf(kl_docx, tmp_dir)
                prefix = f"{monteur['naam']}/" if multi else ""
                zf.write(zzp_docx, f"{prefix}{zzp_fn}.docx")
                zf.write(zzp_pdf,  f"{prefix}{zzp_fn}.pdf")
                zf.write(kl_docx,  f"{prefix}{kl_fn}.docx")
                zf.write(kl_pdf,   f"{prefix}{kl_fn}.pdf")
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)
    return output_zip_path
