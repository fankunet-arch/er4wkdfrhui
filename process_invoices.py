import pdfplumber
import re
import glob
import pandas as pd
import datetime

def parse_mercadona_pdf(filepath):
    items = []
    current_date = None
    current_factura = None
    merchant = 'Mercadona'

    with pdfplumber.open(filepath) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text: continue

            lines = text.split('\n')

            for line in lines:
                if 'MERCADONA S.A. DATOS FISCALES' in line:
                    merchant = 'Mercadona'

                if 'MAKRO' in line.upper():
                    merchant = 'Makro'

                factura_match = re.search(r'Nº Factura:\s*(\S+)\s+Fecha Factura:\s*(\S+)', line)
                if factura_match:
                    current_factura = factura_match.group(1)
                    current_date = factura_match.group(2)

                mercadona_match = re.match(r'^(.*?)\s+([\d]+)\s+([\d,\.]+)\s+([\d,\.]+)\s+(\d+%|\(\*\))\s+([\d,\.]+)\s+([\d,\.]+)$', line)
                if mercadona_match:
                    desc = mercadona_match.group(1).strip()
                    if desc.startswith('PARKING'): continue

                    qty = float(mercadona_match.group(2))

                    def parse_number(num_str):
                        return float(num_str.replace('.', '').replace(',', '.'))

                    p_unit = parse_number(mercadona_match.group(3))
                    b_imp = parse_number(mercadona_match.group(4))
                    iva_rate = mercadona_match.group(5)
                    cuota_iva = parse_number(mercadona_match.group(6))
                    importe = parse_number(mercadona_match.group(7))

                    # Compute con IVA unit price
                    p_unit_con_iva = importe / qty if qty > 0 else 0

                    items.append({
                        '商家': merchant,
                        '发票号码': current_factura,
                        '日期': current_date,
                        '原始名称': desc,
                        '中文名称': '',
                        '数量': qty,
                        '单价金额 (sin IVA)': p_unit,
                        '单品总计金额 (sin IVA)': b_imp,
                        'IVA税率': iva_rate,
                        '单价金额 (con IVA)': round(p_unit_con_iva, 4),
                        '单品总计金额 (con IVA)': importe
                    })
    return items

def parse_makro_excel(filepath):
    df = pd.read_excel(filepath)
    items = []

    for index, row in df.iterrows():
        desc = str(row.get('Descripción Art.', '')).strip()
        if pd.isna(row.get('Descripción Art.')) or desc == 'nan':
            continue

        qty = float(row.get('Cantidad', 0))
        p_unit_sin_iva = float(row.get('Precio/Pieza/Kg', 0))
        total_sin_iva = float(row.get('Precio Neto', 0))
        total_con_iva = float(row.get('Precio Bruto', 0))
        iva_rate = str(row.get('Impuesto', '')).strip()

        merchant = 'Makro'
        if 'Centro ' in row and not pd.isna(row['Centro ']):
            merchant = f"Makro ({row['Centro ']})"

        date = str(row.get('Fecha Factura', '')).strip()
        if hasattr(row.get('Fecha Factura'), 'strftime'):
            date = row['Fecha Factura'].strftime('%d/%m/%Y')
        else:
            # Handle standard formatting for makro date if string
            date = date.replace('-', '/')

        p_unit_con_iva = total_con_iva / qty if qty > 0 else 0

        items.append({
            '商家': merchant,
            '发票号码': str(row.get('Nº Factura', '')).strip(),
            '日期': date,
            '原始名称': desc,
            '中文名称': '',
            '数量': qty,
            '单价金额 (sin IVA)': p_unit_sin_iva,
            '单品总计金额 (sin IVA)': total_sin_iva,
            'IVA税率': iva_rate,
            '单价金额 (con IVA)': round(p_unit_con_iva, 4),
            '单品总计金额 (con IVA)': total_con_iva
        })
    return items

def main():
    all_items = []

    for f in glob.glob('fk/*.pdf'):
        all_items.extend(parse_mercadona_pdf(f))

    for f in glob.glob('fk/*.ods'):
        all_items.extend(parse_makro_excel(f))

    df = pd.DataFrame(all_items)

    # Calculate month from date
    def extract_month(date_str):
        if not date_str: return ''
        # Assume date format could be DD/MM/YYYY or similar based on parsing output
        # E.g. 31/01/2025 or 02-10-2025
        date_str = date_str.replace('-', '/')
        parts = date_str.split('/')
        if len(parts) >= 2:
            return f"{parts[2]}-{parts[1]}" if len(parts[2]) == 4 else f"20{parts[2][-2:]}-{parts[1]}"
        return date_str

    df['月份'] = df['日期'].apply(extract_month)

    # Reorder columns
    columns_order = [
        '商家', '月份', '日期', '发票号码', '原始名称', '中文名称', '数量',
        '单价金额 (sin IVA)', '单品总计金额 (sin IVA)', 'IVA税率',
        '单价金额 (con IVA)', '单品总计金额 (con IVA)'
    ]
    df = df[columns_order]

    output_file = '订货统计.xlsx'
    df.to_excel(output_file, index=False)
    print(f"Data successfully exported to {output_file} ({len(df)} records)")

if __name__ == "__main__":
    main()
