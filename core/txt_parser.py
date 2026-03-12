"""Parser fuer SAP Migration Protokoll-Dateien (TXT/CSV).
Unterstuetzt englische UND deutsche Protokolle.
"""
import os

ACTION_MAP = {
    'Migriert': 'Migrated',
    'Migrated': 'Migrated',
}
STATUS_MAP = {
    'Erfolg': 'Success',
    'Success': 'Success',
}


def parse_migration_protocol(filepath, xml_key_fields=None):
    """Parse Protokoll-Datei.

    Returns:
        headers, data_rows, proto_type ('en' oder 'de')
    """
    with open(filepath, 'r', encoding='utf-8', errors='replace') as f:
        lines = f.read().strip().splitlines()
    if len(lines) < 2:
        raise ValueError(f"Protokoll-Datei ist leer: {os.path.basename(filepath)}")

    headers = [h.strip() for h in lines[0].split('\t')]

    has_product_number = 'Product Number' in headers

    if xml_key_fields is None:
        xml_key_fields = []
    matching_key_fields = [f for f in xml_key_fields if f in headers]

    data_rows = []
    for line in lines[1:]:
        parts = line.split('\t')
        row = {}
        for i, header in enumerate(headers):
            row[header] = parts[i].strip() if i < len(parts) else ''

        raw_action = row.get('Action', '')
        raw_status = row.get('Status', '')
        row['Action_Normalized'] = ACTION_MAP.get(raw_action, raw_action)
        row['Status_Normalized'] = STATUS_MAP.get(raw_status, raw_status)

        if has_product_number:
            row['_match_key'] = row.get('Product Number', '').strip()
        elif matching_key_fields:
            key_parts = [row.get(f, '').strip() for f in matching_key_fields]
            row['_match_key'] = '|'.join(key_parts)
        else:
            non_meta = [h for h in headers if h not in ('Action', 'Status')]
            key_parts = [row.get(h, '').strip() for h in non_meta[:3]]
            row['_match_key'] = '|'.join(key_parts)

        data_rows.append(row)

    proto_type = 'en' if has_product_number else 'de'
    return headers, data_rows, proto_type


def analyze_protocol(data_rows):
    """Analysiere Protokoll - nutzt normalisierte Werte."""
    total = len(data_rows)
    migrated = sum(1 for r in data_rows if r.get('Action_Normalized', '') == 'Migrated')
    success = sum(1 for r in data_rows if r.get('Status_Normalized', '') == 'Success')
    return {
        'total': total,
        'migrated': migrated,
        'not_migrated': total - migrated,
        'success': success,
        'not_success': total - success,
        'all_migrated': migrated == total,
        'all_success': success == total,
    }
