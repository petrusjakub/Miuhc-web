#!/usr/bin/env python3
"""
Extract premium data from xlsx files and generate JavaScript data objects.
Parses both Data_Premi_Lengkap.xlsx (konvensional) and Data_Premi_Syariah_Lengkap.xlsx (syariah).
"""

import zipfile
import xml.etree.ElementTree as ET
import json
import os

NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

# Column mapping: xlsx columns 5-12 (0-indexed: 4-11) map to plan keys
# Columns: Diamond, Ruby, Emerald, Topaz, Topaz ID, Jade, Jade ID, Sapphire
PLAN_COLS = [
    (4, 'diamond'),
    (5, 'ruby'),
    (6, 'emerald'),
    (7, 'topaz'),
    (8, 'topaz_id'),
    (9, 'jade'),
    (10, 'jade_id'),
    (11, 'sapphire'),
]

def parse_xlsx(filepath):
    """Parse an xlsx file and return list of row data (list of lists)."""
    rows_data = []
    with zipfile.ZipFile(filepath, 'r') as z:
        with z.open('xl/worksheets/sheet1.xml') as f:
            tree = ET.parse(f)
            root = tree.getroot()
            rows = root.findall(f'.//{{{NS}}}row')
            for row in rows:
                cells = row.findall(f'{{{NS}}}c')
                row_vals = []
                for cell in cells:
                    is_el = cell.find(f'{{{NS}}}is')
                    if is_el is not None:
                        t_el = is_el.find(f'{{{NS}}}t')
                        row_vals.append(t_el.text if t_el is not None else '')
                    else:
                        v_el = cell.find(f'{{{NS}}}v')
                        row_vals.append(v_el.text if v_el is not None else '')
                rows_data.append(row_vals)
    return rows_data


def parse_value(val):
    """Parse a cell value to int or None."""
    if val is None or val.strip() == '':
        return None
    return int(val)


def build_konv_data(rows):
    """
    Build konvensional data structure.
    Only Pria rows, ages 0-85, 4 rows per age:
      Kamar Normal, Kamar Smart, Rawat Jalan Normal, Rawat Jalan Smart
    """
    data = {
        'kamar': {'normal': {}, 'smart': {}},
        'rawat_jalan': {'normal': {}, 'smart': {}},
    }

    # Initialize plan dicts
    for category in data:
        for typ in data[category]:
            for _, plan_key in PLAN_COLS:
                data[category][typ][plan_key] = {}

    # Skip header row (index 0)
    data_rows = rows[1:]

    for row in data_rows:
        if len(row) < 12:
            continue
        gender = row[0]
        if gender != 'Pria':
            continue
        age = int(row[1])
        kategori = row[2]  # 'Kamar' or 'Rawat Jalan'
        tipe = row[3]      # 'Normal' or 'Smart'

        if kategori == 'Kamar':
            cat_key = 'kamar'
        elif kategori == 'Rawat Jalan':
            cat_key = 'rawat_jalan'
        else:
            continue

        typ_key = tipe.lower()  # 'normal' or 'smart'

        for col_idx, plan_key in PLAN_COLS:
            val = parse_value(row[col_idx]) if col_idx < len(row) else None
            data[cat_key][typ_key][plan_key][age] = val

    return data


def build_syariah_data(rows):
    """
    Build syariah data structure.
    Use Wanita rows, ages 0-79, 6 rows per age:
      Kamar Normal, Kamar Smart, Rawat Jalan Normal, Rawat Jalan Smart,
      Rawat Gigi Normal, Rawat Gigi Smart
    """
    data = {
        'kamar': {'normal': {}, 'smart': {}},
        'rawat_jalan': {'normal': {}, 'smart': {}},
        'rawat_gigi': {'normal': {}, 'smart': {}},
    }

    # Initialize plan dicts
    for category in data:
        for typ in data[category]:
            for _, plan_key in PLAN_COLS:
                data[category][typ][plan_key] = {}

    # Skip header row (index 0)
    data_rows = rows[1:]

    for row in data_rows:
        if len(row) < 12:
            continue
        gender = row[0]
        if gender != 'Wanita':
            continue
        age = int(row[1])
        kategori = row[2]
        tipe = row[3]

        if kategori == 'Kamar':
            cat_key = 'kamar'
        elif kategori == 'Rawat Jalan':
            cat_key = 'rawat_jalan'
        elif kategori == 'Rawat Gigi':
            cat_key = 'rawat_gigi'
        else:
            continue

        typ_key = tipe.lower()

        for col_idx, plan_key in PLAN_COLS:
            val = parse_value(row[col_idx]) if col_idx < len(row) else None
            data[cat_key][typ_key][plan_key][age] = val

    return data


def format_age_dict(d):
    """Format an age dict {0: val, 1: val, ...} as JS object literal."""
    items = []
    for age in sorted(d.keys()):
        val = d[age]
        if val is None:
            items.append(f'{age}:null')
        else:
            items.append(f'{age}:{val}')
    return '{' + ','.join(items) + '}'


def format_data_object(data, var_name):
    """Format the full data object as a JS const declaration."""
    lines = [f'const {var_name} = {{']
    categories = list(data.keys())
    for ci, category in enumerate(categories):
        lines.append(f'  {category}: {{')
        types = list(data[category].keys())
        for ti, typ in enumerate(types):
            lines.append(f'    {typ}: {{')
            plans = list(data[category][typ].keys())
            for pi, plan_key in enumerate(plans):
                age_dict = data[category][typ][plan_key]
                formatted = format_age_dict(age_dict)
                comma = ',' if pi < len(plans) - 1 else ''
                lines.append(f'      {plan_key}: {formatted}{comma}')
            type_comma = ',' if ti < len(types) - 1 else ''
            lines.append(f'    }}{type_comma}')
        cat_comma = ',' if ci < len(categories) - 1 else ''
        lines.append(f'  }}{cat_comma}')
    lines.append('};')
    return '\n'.join(lines)


def main():
    script_dir = os.path.dirname(os.path.abspath(__file__))

    # Parse konvensional
    konv_rows = parse_xlsx(os.path.join(script_dir, 'Data_Premi_Lengkap.xlsx'))
    konv_data = build_konv_data(konv_rows)

    # Parse syariah
    syariah_rows = parse_xlsx(os.path.join(script_dir, 'Data_Premi_Syariah_Lengkap.xlsx'))
    syariah_data = build_syariah_data(syariah_rows)

    # Output JS
    print(format_data_object(konv_data, 'DATA_KONV'))
    print()
    print(format_data_object(syariah_data, 'DATA_SYARIAH'))


if __name__ == '__main__':
    main()
