"""
GTO Overeenkomst Generator
==========================
Vult de 4 originele Word-templates exact in op basis van content-control (sdt) index,
behoudt alle originele opmaak, en exporteert naar Word + PDF in een ZIP.

Templates:
  - nl_zzp:    Vernieuwde_overeenkomst_ZZP_2026_NEDERLAND.docx
  - west_zzp:  Vernieuwde_overeenkomst_ZZP_2026_WEST.docx
  - nl_klant:  NIEUW_projectovereenkomst_KLANT_2026_Nederland.docx
  - west_klant: NIEUW_projectovereenkomst_KLANT_2026_West.docx
"""

import re
import os
import shutil
import zipfile
import subprocess
import io
from pathlib import Path

import tempfile as _tempfile

TEMPLATES_DIR = Path(__file__).parent / "templates"
WORK_DIR = Path(_tempfile.gettempdir()) / "gto_work"
WORK_DIR.mkdir(parents=True, exist_ok=True)

# ── sdt index -> veldnaam mapping per template ─────────────────────────────
# Bepaald door analyse van elk document's content controls in volgorde

SDT_MAP_NL_ZZP = {
    0: "projectnr_combo",      # Projectovereenkomst nummer + persoonsnummer
    2: "monteur_handelsnaam",  # Bedrijfsnaam
    3: "monteur_adres",        # Adres
    4: "monteur_naam",         # naam (vertegenwoordigd door)
    5: "klant_naam",           # KLANT
    6: "project_naam",         # PROJECT NAAM
    7: "project_werkadres",    # PROJECT ADRES
    8: "project_nr",           # PROJECTOVEREENKOMST NR
    10: "monteur_naam2",       # NAAM OPDRACHTNEMER
    11: "monteur_handelsnaam2",# HANDELSNAAM OPDRACHTNEMER
    12: "monteur_kvk",         # KVK-NUMMER OPDRACHTNEMER
    13: "opdrachtomschrijving",# OPDRACHTOMSCHRIJVING
    14: "startdatum",          # STARTDATUM
    16: "uurtarief_zzp",       # UURTARIEF
    18: "handtekendatum",      # Datum: (handtekentabel)
    20: "monteur_handelsnaam3",# Bedrijfsnaam (handtekentabel onderaan)
}

SDT_MAP_WEST_ZZP = {
    0: "projectnr_combo",
    2: "monteur_handelsnaam",
    3: "monteur_adres",
    4: "monteur_naam",
    5: "klant_naam",
    6: "project_naam",
    7: "project_werkadres",
    8: "project_nr",
    10: "monteur_naam2",
    11: "monteur_handelsnaam2",
    12: "monteur_kvk",
    13: "opdrachtomschrijving",
    14: "startdatum",
    16: "uurtarief_zzp",
    18: "handtekendatum",       # Datum (handtekentabel) - West heeft hier WEL inhoud nodig
    20: "monteur_handelsnaam3",
}

SDT_MAP_NL_KLANT = {
    0: "projectnr_combo",       # Projectovereenkomst nummer
    1: "klant_naam",            # "De ondergetekenden: ___, gevestigd aan de"
    2: "klant_adres",           # "___, te dezen rechtsgeldig vertegenwoordigd"
    3: "tekenbevoegde",         # "de heer of mevrouw ___;"
    5: "klant_naam",            # KLANT/OPDRACHTGEVER: (tabel, duplicaat van naam)
    6: "klant_adres",           # ADRES: (tabel, duplicaat van adres)
    7: "klant_kvk",             # KVK-NUMMER
    8: "klant_btw",             # BTW-NUMMER
    9: "klant_contact",         # CONTACTPERSOON
    10: "klant_werkadres",      # WERKADRES
    11: "klant_fmail",          # FACTUUR EMAILADRES OPDRACHTGEVER
    # 12 = betalingstermijn, heeft al vaste waarde "2 weken" - alleen overschrijven indien afwijkend
    14: "monteur_naam",         # NAAM ZELFSTANDIGE
    15: "monteur_handelsnaam",  # HANDELSNAAM ZELFSTANDIGE
    16: "monteur_kvk",          # KVK-NUMMER ZELFSTANDIGE
    17: "opdrachtomschrijving", # OPDRACHTOMSCHRIJVING
    18: "project_nr",           # PROJECTOVEREENKOMST NUMMER
    19: "startdatum",           # STARTDATUM
    20: "einddatum",            # EINDDATUM
    21: "uurtarief_klant",      # UURTARIEF
    24: "tekenbevoegde",        # Akkoord gegaan door Opdrachtgever:
    26: "handtekendatum",       # Datum akkoord gegaan:
}

SDT_MAP_WEST_KLANT = {
    0: "projectnr_combo",       # Projectovereenkomst nummer
    1: "klant_naam",            # "De ondergetekenden: ___, gevestigd aan de"
    2: "klant_adres",           # "___, te dezen rechtsgeldig vertegenwoordigd"
    3: "tekenbevoegde",         # "de heer of mevrouw ___;"
    5: "klant_naam",            # KLANT/OPDRACHTGEVER: (tabel)
    6: "klant_adres",           # ADRES: (tabel)
    7: "klant_kvk",             # KVK-NUMMER
    8: "klant_btw",             # BTW-NUMMER
    9: "klant_contact",         # CONTACTPERSOON
    10: "klant_werkadres",      # WERKADRES
    11: "klant_fmail",          # FACTUUR EMAILADRES OPDRACHTGEVER
    13: "monteur_naam",         # NAAM ZELFSTANDIGE
    14: "monteur_handelsnaam",  # HANDELSNAAM ZELFSTANDIGE
    15: "monteur_kvk",          # KVK-NUMMER ZELFSTANDIGE
    16: "opdrachtomschrijving", # OPDRACHTOMSCHRIJVING
    17: "project_nr",           # PROJECTOVEREENKOMST NUMMER
    18: "startdatum",           # STARTDATUM
    19: "einddatum",            # EINDDATUM
    20: "uurtarief_klant",      # UURTARIEF
    22: "tekenbevoegde",        # Akkoord gegaan door Opdrachtgever:
    24: "handtekendatum",       # Datum akkoord gegaan:
}

TEMPLATE_FILES = {
    "nl_zzp": "Vernieuwde_overeenkomst_ZZP_2026_NEDERLAND.docx",
    "west_zzp": "Vernieuwde_overeenkomst_ZZP_2026_WEST.docx",
    "nl_klant": "NIEUW_projectovereenkomst_KLANT_2026_Nederland.docx",
    "west_klant": "NIEUW_projectovereenkomst_KLANT_2026_West.docx",
}

SDT_MAPS = {
    "nl_zzp": SDT_MAP_NL_ZZP,
    "west_zzp": SDT_MAP_WEST_ZZP,
    "nl_klant": SDT_MAP_NL_KLANT,
    "west_klant": SDT_MAP_WEST_KLANT,
}


def escape_xml(text):
    """Escape tekst voor veilige invoeging in Word XML."""
    if text is None:
        text = ""
    text = str(text)
    text = text.replace('&', '&amp;')
    text = text.replace('<', '&lt;')
    text = text.replace('>', '&gt;')
    return text


def unpack_docx(docx_path, work_dir):
    """
    Eigen, zelfstandige unpack van een .docx (= ZIP archief).
    Geen externe skill-dependency nodig - werkt op elke Python omgeving.
    """
    work_dir = Path(work_dir)
    if work_dir.exists():
        shutil.rmtree(work_dir)
    work_dir.mkdir(parents=True)

    with zipfile.ZipFile(docx_path, "r") as zf:
        zf.extractall(work_dir)

    return work_dir


def pack_docx(work_dir, output_path):
    """
    Eigen, zelfstandige pack van een unpacked directory terug naar .docx.
    Behoudt de originele bestandsvolgorde niet strikt (niet nodig voor geldigheid).
    """
    work_dir = Path(work_dir)
    output_path = Path(output_path)
    if output_path.exists():
        output_path.unlink()

    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zf:
        # [Content_Types].xml moet eerst voor sommige strikte readers; niet strikt vereist maar netjes
        all_files = list(work_dir.rglob("*"))
        all_files = [f for f in all_files if f.is_file()]

        def sort_key(f):
            rel = f.relative_to(work_dir).as_posix()
            return (0 if rel == "[Content_Types].xml" else 1, rel)

        all_files.sort(key=sort_key)

        for f in all_files:
            rel_path = f.relative_to(work_dir).as_posix()
            zf.write(f, rel_path)

    return output_path



    """
    Vind alle TOP-LEVEL <w:sdt>...</w:sdt> blocks met correcte bracket-matching
    (sommige content controls in deze templates zijn 2-3 lagen genest).
    Retourneert lijst van (start, end) tuples voor de buitenste sdt-blocks.
    """
    tag_pattern = re.compile(r'<(/?)w:sdt(\s|>|/>)')
    depth = 0
    top_level_blocks = []
    current_start = None

    for m in tag_pattern.finditer(xml_content):
        is_closing = m.group(1) == '/'
        # Check of het een self-closing achtige <w:sdt/> is (komt hier niet voor, maar safety)
        if not is_closing:
            if depth == 0:
                current_start = m.start()
            depth += 1
        else:
            depth -= 1
            if depth == 0 and current_start is not None:
                end = m.end()
                # vind het echte einde van de closing tag '>'
                close_idx = xml_content.find('>', m.start())
                top_level_blocks.append((current_start, close_idx + 1))
                current_start = None

    return top_level_blocks


def find_balanced_sdt_blocks(xml_content):
    """
    Vind alle TOP-LEVEL <w:sdt>...</w:sdt> blocks met correcte bracket-matching
    (sommige content controls in deze templates zijn 2-3 lagen genest).
    Retourneert lijst van (start, end) tuples voor de buitenste sdt-blocks.
    """
    tag_pattern = re.compile(r'<(/?)w:sdt(\s|>|/>)')
    depth = 0
    top_level_blocks = []
    current_start = None

    for m in tag_pattern.finditer(xml_content):
        is_closing = m.group(1) == '/'
        if not is_closing:
            if depth == 0:
                current_start = m.start()
            depth += 1
        else:
            depth -= 1
            if depth == 0 and current_start is not None:
                end = m.end()
                close_idx = xml_content.find('>', m.start())
                top_level_blocks.append((current_start, close_idx + 1))
                current_start = None

    return top_level_blocks


def fixed_rpr(size_half_points=14):
    """
    Bouw een vaste, schone rPr-XML met Arial in de opgegeven grootte.
    Word gebruikt halve punten: 14 = 7pt (standaard voor ingevulde velden),
    22 = 11pt (voor de bedrijfsnaam onderaan bij ondertekening).
    """
    return (
        f'<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>'
        f'<w:sz w:val="{size_half_points}"/><w:szCs w:val="{size_half_points}"/></w:rPr>'
    )


def fill_sdt_content(xml_content, sdt_index, new_text, blocks_cache=None, font_size=14):
    """
    Vervang de tekst van het TOP-LEVEL sdt-block op de gegeven index (0-based)
    met new_text. Sommige sdt's omvatten een hele <w:tc> (tabelcel) of bevatten
    geneste sdt's - we behouden ALTIJD de volledige structuur (w:tc, w:p, geneste
    w:sdt wrappers) en vervangen alleen de daadwerkelijke <w:r>...<w:t> runs
    binnenin met één run die de nieuwe tekst bevat. Lettertype/grootte wordt
    ALTIJD geforceerd naar Arial in de opgegeven grootte (in halve punten;
    14 = 7pt standaard, 22 = 11pt voor de bedrijfsnaam bij ondertekening),
    ongeacht wat de originele placeholder-stijl was.
    """
    blocks = blocks_cache if blocks_cache is not None else find_balanced_sdt_blocks(xml_content)

    if sdt_index >= len(blocks):
        return xml_content

    start, end = blocks[sdt_index]
    sdt_block = xml_content[start:end]

    escaped_text = escape_xml(new_text)
    needs_preserve = (new_text == '') or (new_text != new_text.strip())
    space_attr = ' xml:space="preserve"' if needs_preserve else ''
    rpr_xml = fixed_rpr(font_size)

    # Vind de buitenste <w:sdtContent>...</w:sdtContent> (balanced, kan geneste sdtContent bevatten)
    content_tag_pattern = re.compile(r'<(/?)w:sdtContent(\s|>|/>)')
    depth = 0
    content_start = None
    content_end = None
    for m in content_tag_pattern.finditer(sdt_block):
        is_closing = m.group(1) == '/'
        if not is_closing:
            if depth == 0:
                content_start = m.end()
                # ga naar het einde van de openings-tag
                content_start = sdt_block.find('>', m.start()) + 1
            depth += 1
        else:
            depth -= 1
            if depth == 0:
                content_end = m.start()
                break

    if content_start is None or content_end is None:
        return xml_content  # onverwachte structuur, sla over

    inner = sdt_block[content_start:content_end]

    # Geval A: de sdtContent bevat direct block-elements (w:tc of w:p) - dit zijn
    # "hele cel" of "hele paragraaf" content controls. We MOETEN die structuur behouden
    # en alleen de tekst diep binnenin vervangen.
    has_block_element = bool(re.search(r'<w:(tc|p)\b', inner))

    if has_block_element:
        # Vervang alle runs (<w:r>...</w:r>) binnen 'inner' door: eerste run krijgt nieuwe tekst
        # met GEFORCEERDE vaste opmaak, overige runs worden volledig verwijderd.
        run_pattern = re.compile(r'<w:r\b[^>]*>(?:(?!</w:r>).)*?</w:r>', re.DOTALL)
        runs = list(run_pattern.finditer(inner))

        if not runs:
            # geen losse runs gevonden, sla over
            return xml_content

        new_first_run = f'<w:r>{rpr_xml}<w:t{space_attr}>{escaped_text}</w:t></w:r>'

        pieces = []
        last_end = 0
        for i, r in enumerate(runs):
            pieces.append(inner[last_end:r.start()])
            if i == 0:
                pieces.append(new_first_run)
            last_end = r.end()
        pieces.append(inner[last_end:])
        new_inner = ''.join(pieces)

        new_sdt_block = sdt_block[:content_start] + new_inner + sdt_block[content_end:]
        return xml_content[:start] + new_sdt_block + xml_content[end:]

    else:
        # Geval B: sdtContent bevat alleen inline content (mogelijk geneste sdt's met simpele runs)
        # Vervang alles met 1 platte run met GEFORCEERDE vaste opmaak.
        new_run = f'<w:r>{rpr_xml}<w:t{space_attr}>{escaped_text}</w:t></w:r>'
        new_sdt_block = sdt_block[:content_start] + new_run + sdt_block[content_end:]
        return xml_content[:start] + new_sdt_block + xml_content[end:]


def strip_placeholder_style(rpr_xml):
    """
    Verwijder de 'Tekstvantijdelijkeaanduiding' (Placeholder Text) rStyle-referentie
    uit een rPr-block. Deze stijl zet tekst grijs (#666666) - bedoeld voor de
    ingebouwde "Klik of tik om tekst in te voeren" hint, maar moet NIET worden
    toegepast op de daadwerkelijk ingevulde data (die moet gewoon zwart zijn).
    """
    return re.sub(r'<w:rStyle\s+w:val="Tekstvantijdelijkeaanduiding"\s*/>', '', rpr_xml)


SIGNATURES_DIR = Path(__file__).parent / "signatures"
SIGNATURE_FILES = {
    "nl": "handtekening_nl.png",
    "west": "handtekening_west.png",
}


def find_signature_image_targets(work_unpack_dir):
    """
    Vind de media-bestanden in word/media/ die als ingebedde handtekening-
    afbeelding dienen (een <w:drawing> met wp:anchor, dus vrij zwevend
    gepositioneerd op specifieke plekken zoals "HANDTEKENING GTO Nederland"
    en de Akkoord-tabel onderaan). Retourneert een set van relatieve image-
    bestandsnamen (bijv. {'image1.png'}) die vervangen moeten worden.
    """
    doc_xml_path = work_unpack_dir / "word" / "document.xml"
    rels_path = work_unpack_dir / "word" / "_rels" / "document.xml.rels"

    with open(doc_xml_path, "r", encoding="utf-8") as f:
        xml_content = f.read()
    with open(rels_path, "r", encoding="utf-8") as f:
        rels_content = f.read()

    # Vind alle r:embed id's die binnen een <wp:anchor> (vrij zwevende afbeelding) voorkomen
    anchor_pattern = re.compile(r'<wp:anchor.*?</wp:anchor>', re.DOTALL)
    embed_ids = set()
    for m in anchor_pattern.finditer(xml_content):
        for rid_match in re.finditer(r'r:embed="(rId\d+)"', m.group()):
            embed_ids.add(rid_match.group(1))

    # Map elke rId naar het bijbehorende media-bestand via de rels-file
    targets = set()
    for rid in embed_ids:
        m = re.search(rf'<Relationship Id="{rid}"[^>]*Target="media/([^"]+)"', rels_content)
        if m:
            targets.add(m.group(1))

    return targets


def inject_signature_image(work_unpack_dir, entiteit):
    """
    Vervang de pixelinhoud van de ingebedde handtekening-afbeelding(en) in het
    template door de juiste, scherpe handtekening voor de gekozen entiteit
    (J. Janssen voor GTO Nederland, Debby van Lith voor GTO West). De originele
    templates verwijzen via een vaste relationship-id naar deze afbeelding op
    twee plekken (onder "HANDTEKENING GTO Nederland/West" en in de "Akkoord"
    handtekentabel onderaan) - door alleen de bytes van het bestand te vervangen
    (niet de XML-structuur), verschijnt de juiste handtekening automatisch op
    beide plekken zonder dat we de document-structuur hoeven aan te passen.
    """
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


def find_balanced_tags(xml_content, tagname):
    """
    Generieke balanced bracket-matcher voor een willekeurige w:-tag (bijv. 'tc', 'tr', 'tbl', 'p').
    Houdt correct rekening met zelfsluitende varianten zoals <w:p .../> (komt voor bij
    volledig lege paragrafen zonder pPr/inhoud) - die worden als losstaand blok geteld,
    niet als "open zonder close" wat de depth-tracking zou verstoren.
    """
    tag_pattern = re.compile(rf'<(/?)w:{tagname}(?=[\s>/])[^>]*?(/?)>')
    depth = 0
    blocks = []
    current_start = None
    for m in tag_pattern.finditer(xml_content):
        is_closing = m.group(1) == '/'
        is_selfclosing = m.group(2) == '/'
        if is_closing:
            depth -= 1
            if depth == 0 and current_start is not None:
                blocks.append((current_start, m.end()))
                current_start = None
        elif is_selfclosing:
            if depth == 0:
                blocks.append((m.start(), m.end()))
            # zelfsluitend op diepere niveaus heeft geen effect op depth
        else:
            if depth == 0:
                current_start = m.start()
            depth += 1
    return blocks


def find_balanced_paragraphs(xml_content):
    """Vind alle TOP-LEVEL <w:p>...</w:p> of <w:p .../> blocks."""
    return find_balanced_tags(xml_content, "p")


def collapse_empty_paragraph_runs(xml_content, max_consecutive=0):
    """
    Verwijder reeksen van 2+ opeenvolgende, volledig lege <w:p> paragrafen
    (geen tekst, geen sdt, geen drawing/afbeelding, geen tabel, en NIET de
    enige paragraaf binnen een tabelcel - Word vereist altijd minstens één
    paragraaf per cel) tot maximaal max_consecutive exemplaren. Dit fixt de
    storende witruimtes tussen artikelen in de klant-templates die in het
    origineel document zaten. Gebruikt balanced bracket-matching om gegarandeerd
    geen tabelcel-paragrafen of content-controls per ongeluk te beschadigen.

    KRITIEK: een verwijderbare reeks mag NOOIT een tabelgrens (</w:tc>, </w:tr>,
    </w:tbl>) overschrijden, ook niet als de paragrafen aan weerszijden allebei
    op zichzelf verwijderbaar zouden zijn - anders kan een hele tabelcel zijn
    enige overgebleven paragraaf verliezen wat het document corrumpeert.
    """
    blocks = find_balanced_paragraphs(xml_content)

    # Bepaal voor elke paragraaf bij welke "tabel-context" hij hoort, zodat we
    # nooit een reeks vormen die twee verschillende contexten overspant.
    # We gebruiken de positie van de dichtstbijzijnde tabel-gerelateerde sluit-tag
    # (</w:tc>, </w:tbl>) die na de paragraaf komt als "context grens-marker".
    boundary_pattern = re.compile(r'</w:tc>|</w:tbl>|</w:tr>')
    boundaries = [m.start() for m in boundary_pattern.finditer(xml_content)]

    def context_id(p_start):
        # Het aantal boundary-markers vóór deze paragraaf bepaalt zijn "context" -
        # twee paragrafen met hetzelfde aantal boundaries ervoor zitten in hetzelfde
        # tabelcel/buiten-tabel segment.
        count = 0
        for b in boundaries:
            if b < p_start:
                count += 1
            else:
                break
        return count

    # Bepaal welke <w:p> blocks zich bevinden binnen een tabelcel (<w:tc>...</w:tc>)
    # die GEEN andere niet-lege paragraaf overhoudt - die paragraaf is dan verplicht.
    tc_blocks = find_balanced_tags(xml_content, "tc")
    protected_starts = set()
    for tc_start, tc_end in tc_blocks:
        tc_text = xml_content[tc_start:tc_end]
        p_count_in_cell = len(re.findall(r'<w:p\b', tc_text))
        if p_count_in_cell <= 1:
            for s, e in blocks:
                if tc_start < s < tc_end:
                    protected_starts.add(s)

    def is_safe_to_remove(p_text, p_start):
        if p_start in protected_starts:
            return False
        texts = re.findall(r'<w:t[^>]*>([^<]*)</w:t>', p_text)
        if ''.join(texts).strip():
            return False
        if '<w:sdt' in p_text or '<w:drawing' in p_text or '<w:tbl' in p_text:
            return False
        return True

    to_remove_spans = []
    i = 0
    while i < len(blocks):
        s, e = blocks[i]
        p_text = xml_content[s:e]
        if is_safe_to_remove(p_text, s):
            ctx = context_id(s)
            j = i
            while j + 1 < len(blocks):
                s2, e2 = blocks[j + 1]
                if is_safe_to_remove(xml_content[s2:e2], s2) and context_id(s2) == ctx:
                    j += 1
                else:
                    break
            run_length = j - i + 1
            if run_length >= 2:
                keep_until = i + max_consecutive
                remove_start = blocks[keep_until][0]
                remove_end = blocks[j][1]
                to_remove_spans.append((remove_start, remove_end))
            i = j + 1
        else:
            i += 1

    # Verwijder van achter naar voor zodat eerdere posities niet verschuiven
    for start, end in reversed(to_remove_spans):
        xml_content = xml_content[:start] + xml_content[end:]

    return xml_content


def generate_docx(template_key, data, output_path):
    """
    Genereer een ingevuld .docx bestand op basis van de template en data dict.
    data bevat alle veldnamen zoals gedefinieerd in de SDT_MAP_* dicts.
    """
    template_path = TEMPLATES_DIR / TEMPLATE_FILES[template_key]
    sdt_map = SDT_MAPS[template_key]

    work_unpack = WORK_DIR / f"unpack_{template_key}_{os.getpid()}_{id(data)}"

    # Eigen, zelfstandige unpack (geen externe skill-dependency)
    unpack_docx(template_path, work_unpack)

    doc_xml_path = work_unpack / "word" / "document.xml"
    with open(doc_xml_path, "r", encoding="utf-8") as f:
        xml_content = f.read()

    # Vervang de ingebedde handtekening-afbeelding(en) in het template door de
    # juiste, scherpe handtekening voor deze entiteit (J. Janssen voor GTO
    # Nederland, Debby van Lith voor GTO West). Werkt op de losse bestanden
    # in word/media/, dus VOOR het inlezen van xml_content hierboven zou ook
    # gekund hebben, maar de afbeelding-vervanging raakt de XML niet aan.
    entiteit_voor_handtekening = "west" if template_key.startswith("west") else "nl"
    inject_signature_image(work_unpack, entiteit_voor_handtekening)

    # LET OP: de witruimte-fix (collapse_empty_paragraph_runs) is BEWUST uitgeschakeld.
    # De eerste tabel in de klant-templates is "floating" gepositioneerd
    # (w:tblpPr met vertAnchor="text"), verankerd aan een specifieke paragraaf.
    # Het verwijderen van paragrafen elders in het document verstoort dit anker
    # waardoor de tabel onzichtbaar buiten het canvas terechtkomt in LibreOffice -
    # een ernstiger regressie dan de oorspronkelijke witruimte. Tot een veiligere
    # aanpak is gevonden die het anker behoudt, laten we dit artefact ongemoeid.
    # if template_key in ("nl_klant", "west_klant"):
    #     xml_content = collapse_empty_paragraph_runs(xml_content, max_consecutive=0)

    # Vul elk sdt-veld in op basis van de mapping.
    # Vaste lettertype-regel: Arial 7pt voor alle ingevulde velden, BEHALVE de
    # bedrijfsnaam-velden in de handtekentabel onderaan (monteur_handelsnaam3 /
    # de "Bedrijfsnaam" rij bij Akkoord ZZP'er) die Arial 11pt moeten zijn.
    GROTE_FONT_VELDEN = {"monteur_handelsnaam3"}

    for idx in sorted(sdt_map.keys(), reverse=True):
        field_name = sdt_map[idx]
        value = data.get(field_name, "")
        font_size = 22 if field_name in GROTE_FONT_VELDEN else 14
        blocks = find_balanced_sdt_blocks(xml_content)
        xml_content = fill_sdt_content(xml_content, idx, value, blocks_cache=blocks, font_size=font_size)

    with open(doc_xml_path, "w", encoding="utf-8") as f:
        f.write(xml_content)

    # Fix: verwijder de broken externe attachedTemplate-verwijzing die in de originele
    # GTO-templates zit (verwijst naar een lokaal pad op een collega's computer).
    # Dit is een onzichtbare Word-instelling, geen zichtbare inhoud.
    settings_rels_path = work_unpack / "word" / "_rels" / "settings.xml.rels"
    settings_path = work_unpack / "word" / "settings.xml"
    if settings_rels_path.exists():
        with open(settings_rels_path, "r", encoding="utf-8") as f:
            rels_content = f.read()
        if "attachedTemplate" in rels_content:
            rels_content = re.sub(
                r'<Relationship[^>]*attachedTemplate[^>]*/>',
                '',
                rels_content
            )
            with open(settings_rels_path, "w", encoding="utf-8") as f:
                f.write(rels_content)
    if settings_path.exists():
        with open(settings_path, "r", encoding="utf-8") as f:
            settings_content = f.read()
        if "attachedTemplate" in settings_content:
            settings_content = re.sub(
                r'<w:attachedTemplate[^/]*/>',
                '',
                settings_content
            )
            with open(settings_path, "w", encoding="utf-8") as f:
                f.write(settings_content)

    # Eigen, zelfstandige pack (geen externe skill-dependency)
    pack_docx(work_unpack, output_path)

    shutil.rmtree(work_unpack, ignore_errors=True)
    return output_path


def docx_to_pdf(docx_path, output_dir):
    """Converteer docx naar pdf via LibreOffice (rechtstreeks, geen skill-wrapper nodig)."""
    subprocess.run(
        ["soffice", "--headless", "--convert-to", "pdf",
         "--outdir", str(output_dir), str(docx_path)],
        check=True, capture_output=True, timeout=60
    )
    pdf_path = Path(output_dir) / (Path(docx_path).stem + ".pdf")
    return pdf_path


def format_kvk_or_regon(kvk):
    """
    Bepaal automatisch of een registratienummer een KvK-nummer of een
    Pools REGON-nummer is, en zet 'REGON' ervoor als dat het geval is.

    KvK-nummer:   8 cijfers (Nederland)
    REGON-nummer: 9 cijfers (Poolse rechtspersoon) of 14 cijfers (Pools filiaal)

    Als het nummer al met 'REGON' begint wordt het ongewijzigd teruggegeven.
    Als het geen zuiver getal is (bijv. al een prefix zoals 'KvK') ook ongewijzigd.
    """
    if not kvk:
        return kvk
    s = str(kvk).strip()
    if s.upper().startswith('REGON'):
        return s  # al correct gelabeld
    digits_only = ''.join(c for c in s if c.isdigit())
    if len(digits_only) in (9, 14) and digits_only == s.replace(' ', ''):
        return f"REGON {digits_only}"
    return s
    """
    Formatteer het uurtarief veld: altijd 2 decimalen, plus optioneel reisuur erachter.

    De ZZP-templates hebben het euroteken al vast staan in het document
    ("UURTARIEF (BTW verlegd): €  ___") - daar gebruiken we met_euroteken=False.
    De KLANT-templates hebben dat NIET vast staan - daar gebruiken we
    met_euroteken=True zodat het € teken wel verschijnt.

    Input kan zijn: '45', '45,00', '45.00', '€ 45,00' etc.
    """
    if not tarief or not str(tarief).strip():
        return ""

    raw = str(tarief).strip()
    raw = raw.replace("€", "").strip()
    raw = raw.replace(".", ",")

    try:
        numeric = float(raw.replace(",", "."))
        formatted = f"{numeric:.2f}".replace(".", ",")
    except ValueError:
        formatted = raw

    if met_euroteken:
        formatted = f"€ {formatted}"

    if reisuur_tekst and reisuur_tekst.strip():
        return f"{formatted} + {reisuur_tekst.strip()}"
    return formatted


def build_filename(doc_type, entiteit, startdatum_iso, persnr, klantnaam, projectnr):
    """
    Bouw de bestandsnaam volgens GTO-conventie.
    doc_type: 'zzp' of 'klant'
    entiteit: 'nl' of 'west'
    startdatum_iso: 'JJJJ-MM-DD'
    persnr: bijv '283' of 'W133' (al met W-prefix indien west)
    """
    klantnaam_clean = re.sub(r'[\\/:*?"<>|]', '', klantnaam).strip()

    if doc_type == "zzp":
        return f"{startdatum_iso} PO {persnr} - {klantnaam_clean} {projectnr}"
    else:
        return f"{startdatum_iso} PO {klantnaam_clean} - {persnr} {projectnr}"


def build_field_data(entiteit, monteur, klant, project, tekenbevoegde,
                      startdatum_nl, handtekendatum_nl, uurtarief_zzp, uurtarief_klant,
                      reisuur_tekst, opdrachtomschrijving):
    """
    Bouw de complete data-dict voor zowel ZZP als Klant template, op basis van
    de profielen uit het platform. Datums hier in NL-formaat (DD-MM-JJJJ) zoals
    de overeenkomst ze toont.

    monteur: dict met naam, handelsnaam, kvk, adres, persnr (al met W-prefix indien west)
    klant: dict met naam, adres, kvk, btw, contact, fmail
    project: dict met nr, naam, werkadres
    tekenbevoegde: naam van de tekenbevoegde voor dit project
    """
    persnr = monteur["persnr"]
    projectnr_combo = f"{project['nr']} + {persnr}"

    common = {
        "projectnr_combo": projectnr_combo,
        "klant_naam": klant["naam"],
        "klant_adres": klant["adres"],
        "klant_kvk": klant["kvk"],
        "klant_btw": klant.get("btw", ""),
        "klant_contact": klant.get("contact", ""),
        "klant_werkadres": project.get("werkadres", klant["adres"]),
        "klant_fmail": klant.get("fmail", ""),
        "tekenbevoegde": tekenbevoegde,
        "monteur_naam": monteur["naam"],
        "monteur_handelsnaam": monteur["handelsnaam"],
        "monteur_adres": monteur.get("adres", ""),
        "monteur_kvk": format_kvk_or_regon(monteur["kvk"]),
        "opdrachtomschrijving": opdrachtomschrijving,
        "project_nr": project["nr"],
        "project_naam": project.get("naam", ""),
        "project_werkadres": project.get("werkadres", ""),
        "startdatum": startdatum_nl,
        "handtekendatum": handtekendatum_nl,
        "einddatum": "3 tot 6 maanden",
    }

    zzp_data = {
        **common,
        "monteur_naam2": monteur["naam"],
        "monteur_handelsnaam2": monteur["handelsnaam"],
        "monteur_handelsnaam3": monteur["handelsnaam"],
        "uurtarief_zzp": format_uurtarief(uurtarief_zzp, reisuur_tekst, met_euroteken=False),
    }

    klant_data = {
        **common,
        "uurtarief_klant": format_uurtarief(uurtarief_klant, reisuur_tekst, met_euroteken=True),
    }

    return zzp_data, klant_data


REQUIRED_FIELDS_MELDING = {
    "tekenbevoegde": "Tekenbevoegde van de klant",
    "klant.naam": "Klantnaam",
    "klant.adres": "Klantadres",
    "klant.kvk": "KvK-nummer klant",
    "project.nr": "Projectnummer",
    "project.naam": "Projectnaam",
    "project.werkadres": "Werkadres project",
    "monteur.naam": "Naam monteur",
    "monteur.handelsnaam": "Handelsnaam monteur",
    "monteur.kvk": "KvK-nummer monteur",
    "monteur.persnr": "Persoonsnummer monteur",
}


def validate_required_fields(entiteit, monteurs, klant, project, tekenbevoegde, opdrachtomschrijving):
    """
    Controleer of alle velden aanwezig zijn die nodig zijn om een volledig
    ingevulde overeenkomst te genereren. Retourneert een lijst van ontbrekende
    velden (lege dicts/strings) zodat het platform de gebruiker gericht om
    aanvulling kan vragen voordat er wordt gegenereerd, in plaats van een
    overeenkomst met lege plekken op te leveren.

    Retourneert: lijst van dicts {"veld": "...", "label": "...", "context": "..."}
    """
    missing = []

    if not tekenbevoegde or not tekenbevoegde.strip():
        missing.append({"veld": "tekenbevoegde", "label": "Tekenbevoegde van de klant", "context": klant.get("naam", "")})

    for key, label in [("naam", "Klantnaam"), ("adres", "Klantadres"), ("kvk", "KvK-nummer klant")]:
        if not klant.get(key, "").strip():
            missing.append({"veld": f"klant.{key}", "label": label, "context": klant.get("naam", "")})

    for key, label in [("nr", "Projectnummer"), ("naam", "Projectnaam"), ("werkadres", "Werkadres project")]:
        if not project.get(key, "").strip():
            missing.append({"veld": f"project.{key}", "label": label, "context": project.get("naam", project.get("nr", ""))})

    if not opdrachtomschrijving or not opdrachtomschrijving.strip():
        missing.append({"veld": "opdrachtomschrijving", "label": "Opdrachtomschrijving", "context": project.get("naam", "")})

    for m in monteurs:
        mnaam = m.get("naam", "(naam onbekend)")
        for key, label in [("naam", "Naam"), ("handelsnaam", "Handelsnaam"), ("kvk", "KvK-nummer")]:
            if not m.get(key, "").strip():
                missing.append({"veld": f"monteur.{key}", "label": f"{label} monteur", "context": mnaam, "monteur_id": m.get("id")})
        if not m.get("persnr", "").strip():
            entiteit_label = "GTO West" if entiteit == "west" else "GTO Nederland"
            missing.append({"veld": "monteur.persnr", "label": f"Persoonsnummer ({entiteit_label})", "context": mnaam, "monteur_id": m.get("id")})

    return missing


def generate_full_package(entiteit, monteurs, klant, project, tekenbevoegde,
                           startdatum_nl, startdatum_iso, handtekendatum_nl,
                           tarieven_per_monteur, opdrachtomschrijving, output_zip_path):
    """
    Genereer de volledige ZIP met Word + PDF voor alle gekozen monteurs.

    entiteit: 'nl' of 'west'
    monteurs: lijst van monteur-dicts
    tarieven_per_monteur: dict {monteur_id: {'zzp': '45,00', 'klant': '54,00', 'reisuur': '1 reisuur 1:1'}}
    output_zip_path: pad waar de ZIP wordt weggeschreven

    Retourneert: pad naar de gegenereerde ZIP
    """
    template_zzp_key = f"{entiteit}_zzp"
    template_klant_key = f"{entiteit}_klant"

    tmp_dir = WORK_DIR / f"package_{os.getpid()}_{id(monteurs)}"
    tmp_dir.mkdir(parents=True, exist_ok=True)

    multi = len(monteurs) > 1

    with zipfile.ZipFile(output_zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for monteur in monteurs:
            tarieven = tarieven_per_monteur.get(monteur["id"], {})
            zzp_data, klant_data = build_field_data(
                entiteit=entiteit,
                monteur=monteur,
                klant=klant,
                project=project,
                tekenbevoegde=tekenbevoegde,
                startdatum_nl=startdatum_nl,
                handtekendatum_nl=handtekendatum_nl,
                uurtarief_zzp=tarieven.get("zzp", ""),
                uurtarief_klant=tarieven.get("klant", ""),
                reisuur_tekst=tarieven.get("reisuur", ""),
                opdrachtomschrijving=opdrachtomschrijving,
            )

            persnr = monteur["persnr"]
            zzp_filename = build_filename("zzp", entiteit, startdatum_iso, persnr, klant["naam"], project["nr"])
            klant_filename = build_filename("klant", entiteit, startdatum_iso, persnr, klant["naam"], project["nr"])

            zzp_docx_path = tmp_dir / f"{zzp_filename}.docx"
            klant_docx_path = tmp_dir / f"{klant_filename}.docx"

            generate_docx(template_zzp_key, zzp_data, zzp_docx_path)
            generate_docx(template_klant_key, klant_data, klant_docx_path)

            zzp_pdf_path = docx_to_pdf(zzp_docx_path, tmp_dir)
            klant_pdf_path = docx_to_pdf(klant_docx_path, tmp_dir)

            # Bepaal het pad binnen de ZIP: submap per monteur als er meerdere zijn
            prefix = f"{monteur['naam']}/" if multi else ""

            zf.write(zzp_docx_path, f"{prefix}{zzp_filename}.docx")
            zf.write(zzp_pdf_path, f"{prefix}{zzp_filename}.pdf")
            zf.write(klant_docx_path, f"{prefix}{klant_filename}.docx")
            zf.write(klant_pdf_path, f"{prefix}{klant_filename}.pdf")

    shutil.rmtree(tmp_dir, ignore_errors=True)
    return output_zip_path
