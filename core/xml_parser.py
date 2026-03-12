"""XML-Parser fuer SAP Migration Checker SpreadsheetML Dateien.
Unterstuetzt englische UND deutsche SAP-XMLs.
"""
import xml.etree.ElementTree as ET
import os
import re

NS = {
    'ss': 'urn:schemas-microsoft-com:office:spreadsheet',
    'o': 'urn:schemas-microsoft-com:office:office',
}

# Mapping: Englisch + Deutsch
FIELD_LIST_NAMES = ['Field List', 'Feldliste']
KEY_GROUP_NAMES = ['Key', 'Schlüssel']
MANDATORY_MARKERS = ['(mandatory)', '(obligatorisch)']


def clean_xml_content(filepath):
    with open(filepath, 'r', encoding='utf-8', errors='replace') as f:
        content = f.read()
    content = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]', '', content)
    return content


def parse_xml_safe(filepath):
    try:
        tree = ET.parse(filepath)
        return tree.getroot()
    except ET.ParseError:
        try:
            content = clean_xml_content(filepath)
            return ET.fromstring(content)
        except Exception as e:
            raise ValueError(
                f"XML konnte nicht gelesen werden: {os.path.basename(filepath)}\n"
                f"Details: {str(e)}"
            )


def parse_row(row):
    cells = row.findall('ss:Cell', NS)
    row_data = {}
    current_col = 1
    for cell in cells:
        idx = cell.get('{urn:schemas-microsoft-com:office:spreadsheet}Index')
        if idx:
            current_col = int(idx)
        data = cell.find('ss:Data', NS)
        value = data.text.strip() if (data is not None and data.text) else ''
        row_data[current_col] = value
        current_col += 1
    return row_data


def _is_field_list_sheet(name):
    """Prueft ob Sheet 'Field List' oder 'Feldliste' heisst."""
    return name in FIELD_LIST_NAMES


def _is_key_group(group_name):
    """Prueft ob Group 'Key' oder 'Schluessel' heisst."""
    return group_name in KEY_GROUP_NAMES


def _extract_mandatory_name(text):
    """Extrahiert Sheet-Name aus 'Name (mandatory)' oder 'Name (obligatorisch)'."""
    for marker in MANDATORY_MARKERS:
        if marker in text:
            return text.replace(marker, '').strip()
    return None


def find_first_mandatory_sheet(root):
    for ws in root.findall('.//ss:Worksheet', NS):
        name = ws.get('{urn:schemas-microsoft-com:office:spreadsheet}Name')
        if not _is_field_list_sheet(name):
            continue

        table = ws.find('ss:Table', NS)
        rows = table.findall('ss:Row', NS)
        first_mandatory_name = None
        key_fields = []
        key_tech_names = []
        in_key_section = False

        for row in rows:
            rd = parse_row(row)
            values = list(rd.values())

            if first_mandatory_name is None:
                for v in values:
                    extracted = _extract_mandatory_name(str(v))
                    if extracted:
                        first_mandatory_name = extracted
                        break
                continue

            group_name = rd.get(3, '')
            field_desc = rd.get(4, '')
            sap_field = rd.get(10, '')

            # Neues Sheet in Spalte 2? -> Stopp nach erstem Sheet!
            sheet_col = rd.get(2, '')
            if sheet_col.strip() and key_fields:
                break

            if _is_key_group(group_name):
                in_key_section = True
                if field_desc:
                    key_fields.append(field_desc)
                    if sap_field:
                        key_tech_names.append(sap_field)
                continue

            if in_key_section:
                if group_name and not _is_key_group(group_name):
                    break
                if field_desc:
                    key_fields.append(field_desc)
                    if sap_field:
                        key_tech_names.append(sap_field)

        return first_mandatory_name, key_fields, key_tech_names

    return None, [], []


def extract_key_data(filepath):
    root = parse_xml_safe(filepath)
    sheet_name, key_fields, key_tech_names = find_first_mandatory_sheet(root)

    if not sheet_name or not key_tech_names:
        raise ValueError(
            f"Kein mandatory/obligatorisch Sheet oder Key/Schluessel-Felder "
            f"in {os.path.basename(filepath)} gefunden!"
        )

    target_ws = None
    for ws in root.findall('.//ss:Worksheet', NS):
        ws_name = ws.get('{urn:schemas-microsoft-com:office:spreadsheet}Name')
        if ws_name == sheet_name:
            target_ws = ws
            break

    if target_ws is None:
        raise ValueError(f"Sheet '{sheet_name}' nicht in {os.path.basename(filepath)} gefunden!")

    table = target_ws.find('ss:Table', NS)
    rows = table.findall('ss:Row', NS)

    # Header suchen - flexibel (nicht fix Zeile 5)
    header_row = None
    header_idx = None
    for idx in range(min(10, len(rows))):
        rd = parse_row(rows[idx])
        vals = set(rd.values())
        matches = sum(1 for tn in key_tech_names if tn in vals)
        if matches >= len(key_tech_names) * 0.5:
            header_row = rd
            header_idx = idx
            break

    if header_row is None:
        header_row = parse_row(rows[4])
        header_idx = 4

    col_map = {tech_name: col_nr for col_nr, tech_name in header_row.items()}

    key_col_indices = []
    for tech_name in key_tech_names:
        if tech_name in col_map:
            key_col_indices.append(col_map[tech_name])
        else:
            raise ValueError(f"Feld '{tech_name}' nicht im Sheet '{sheet_name}' gefunden!")

    data_start = header_idx + 4
    data_rows = []
    for row in rows[data_start:]:
        rd = parse_row(row)
        if not rd:
            continue
        row_values = {}
        for field_name, col_idx in zip(key_fields, key_col_indices):
            row_values[field_name] = rd.get(col_idx, '')
        data_rows.append(row_values)

    return tuple(key_fields), key_fields, data_rows
