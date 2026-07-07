#!/usr/bin/env python3
"""
NetCom BW Export Script
Generates three output files from Update_GPG_*.xlsx:
  - *-Standorte.csv   (Locations import)
  - *-Verträge.csv    (Contracts import)
  - *-Namen.xlsx      (Name parsing review)
"""

import sys, glob, os
import pandas as pd
from process_excel import parse_street_address, parse_name


def fmt_date(val):
    if pd.isnull(val):
        return ''
    try:
        return pd.Timestamp(val).strftime('%d.%m.%y')
    except Exception:
        return str(val)


def fmt_int(val):
    if pd.isnull(val):
        return ''
    try:
        return int(val)
    except Exception:
        return val


def process_file(input_path, output_dir):
    df = pd.read_excel(input_path, dtype={'Ansprechpartner Telefon': str})
    stem = os.path.splitext(os.path.basename(input_path))[0]

    rows_s, rows_v, rows_n = [], [], []

    for _, row in df.iterrows():
        s = parse_street_address(
            str(row['Anschlussadresse Straße']) if pd.notna(row.get('Anschlussadresse Straße')) else ''
        )
        n = parse_name(
            str(row['Kundename']) if pd.notna(row.get('Kundename')) else ''
        )
        zip_code = row.get('Anschlusadresse PLZ', '')
        city     = row.get('Anschlusadresse Ort', '')
        city_p   = row.get('Anschlusadresse Ortsteil', '')
        we       = fmt_int(row.get('Wohneinheiten'))
        phone    = '' if pd.isnull(row['Ansprechpartner Telefon']) else str(row['Ansprechpartner Telefon']).strip()

        rows_s.append({
            'locations_no':       '',
            'lie_zip':            zip_code,
            'lie_city':           city,
            'lie_city_part':      city_p,
            'lie_street':         s.get('street_name', ''),
            'lie_housenoonly':    s.get('house_number', ''),
            'lie_housenoonlyext': s.get('house_number_suffix', ''),
            'wohneinheiten':      we,
            'flurstueck':         '',
        })
        rows_v.append({
            'busclient':          'NetCom BW',
            'intkundennummer':    int(row['Kundennummer']),
            'firstname':          n.get('first_name', ''),
            'lastname':           n.get('last_name', ''),
            'mailingzip':         zip_code,
            'mailingcity':        city,
            'mailingcitypart':    city_p,
            'mailingstreet':      s.get('street_name', ''),
            'mailinghousenr':     s.get('house_number', ''),
            'mailinghousenrpart': s.get('house_number_suffix', ''),
            'intvertragsnummer':  int(row['Vertragsnummer']),
            'contstatus':         row.get('Status', ''),
            'contdateadd':        fmt_date(row.get('Erstellt am')),
            'contdatesigned':     fmt_date(row.get('Datum der Auftragsstellung (Unterschrift)')),
            'homephone':          phone,
            'email':              str(row['Ansprechpartner E-Mail']) if pd.notna(row.get('Ansprechpartner E-Mail')) else '',
            'wohn_einheiten':     we,
            'plotno':             '',
        })
        rows_n.append({
            'Kundennummer': row['Kundennummer'],
            'Kundename':    str(row['Kundename']),
            'Vorname':      n.get('first_name', ''),
            'Nachname':     n.get('last_name', ''),
            'Typ':          'Organisation' if n.get('name_type') == 'ORGANIZATION' else 'Person',
        })

    os.makedirs(output_dir, exist_ok=True)

    path_s = os.path.join(output_dir, f'{stem}-Standorte.csv')
    path_v = os.path.join(output_dir, f'{stem}-Verträge.csv')
    path_n = os.path.join(output_dir, f'{stem}-Namen.xlsx')

    pd.DataFrame(rows_s).to_csv(path_s, sep=';', index=False, encoding='utf-8')
    pd.DataFrame(rows_v).to_csv(path_v, sep=';', index=False, encoding='utf-8')
    pd.DataFrame(rows_n).to_excel(path_n, index=False)

    print(f"  ✅ {os.path.basename(path_s)}")
    print(f"  ✅ {os.path.basename(path_v)}")
    print(f"  ✅ {os.path.basename(path_n)}")

    return pd.DataFrame(rows_n)


def main():
    input_dir  = 'input/NetCom_BW'
    output_dir = 'output/NetCom_BW'

    files = sorted(glob.glob(os.path.join(input_dir, 'Update_GPG_*.xlsx')))
    if not files:
        print(f'Keine Update_GPG_*.xlsx Dateien in {input_dir} gefunden.')
        sys.exit(1)

    # Process the most recent file (highest timestamp in filename)
    input_path = files[-1]
    print(f'\nVerarbeite: {os.path.basename(input_path)}')
    namen = process_file(input_path, output_dir)
    print('\nNamen-Auswertung:')
    print(namen.to_string(index=False))


if __name__ == '__main__':
    main()
