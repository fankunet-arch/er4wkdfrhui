import pdfplumber
import re
import glob
import pandas as pd
import datetime
import os
import time
from deep_translator import GoogleTranslator

def categorize_item(desc):
    desc_upper = desc.upper()

    # Salmon categorization based on rules
    if 'SALMON' in desc_upper:
        if '7-8' in desc_upper:
            return '三文鱼 7-8kg'
        elif '5-6' in desc_upper:
            return '三文鱼 5-6kg'
        elif '8-9' in desc_upper:
            return '三文鱼 8-9kg'
        else:
            return '三文鱼 (其他)'

    # Chicken / Pollo / Gallina categorization
    if 'MEDIA GALLINA' in desc_upper:
        return '半只母鸡'
    elif 'GALLINA' in desc_upper:
        return '母鸡'

    if 'ALBÓNDIGAS POLLO' in desc_upper or 'ALBONDIGAS POLLO' in desc_upper:
        return '鸡肉丸'

    if 'ALAS' in desc_upper or 'ALITA' in desc_upper or 'ALON' in desc_upper or 'ALÓN' in desc_upper:
        if 'POLLO' in desc_upper or 'CONG' in desc_upper or 'PARTIDAS' in desc_upper:
            return '鸡翅'

    if 'PATAS POLLO' in desc_upper or 'PATAS DE POLLO' in desc_upper:
        return '鸡爪'

    if 'CONTRAMUSLO' in desc_upper:
        return '去骨鸡腿肉 (Contramuslo)'

    if 'PECHUGA' in desc_upper:
        return '鸡胸肉 (Pechuga)'

    # If no specific rule matched, return a generic translated placeholder or None
    return None

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

                    p_unit_con_iva = importe / qty if qty > 0 else 0

                    items.append({
                        '商家': merchant,
                        '发票号码': current_factura,
                        '日期': current_date,
                        '原始名称': desc,
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
            date = date.replace('-', '/')

        p_unit_con_iva = total_con_iva / qty if qty > 0 else 0

        items.append({
            '商家': merchant,
            '发票号码': str(row.get('Nº Factura', '')).strip(),
            '日期': date,
            '原始名称': desc,
            '数量': qty,
            '单价金额 (sin IVA)': p_unit_sin_iva,
            '单品总计金额 (sin IVA)': total_sin_iva,
            'IVA税率': iva_rate,
            '单价金额 (con IVA)': round(p_unit_con_iva, 4),
            '单品总计金额 (con IVA)': total_con_iva
        })
    return items

def translate_descriptions(unique_desc):
    translations = {}
    print(f"Translating {len(unique_desc)} unique descriptions...")
    translator = GoogleTranslator(source='es', target='zh-CN')
    for i, desc in enumerate(unique_desc):
        if not desc:
            translations[desc] = ""
            continue

        try:
            translated = translator.translate(desc)
            translations[desc] = translated
        except Exception as e:
            print(f"Error translating '{desc}': {e}")
            translations[desc] = desc # Fallback to original
            time.sleep(1) # Extra delay on error

        if (i+1) % 50 == 0:
            print(f"Translated {i+1}/{len(unique_desc)} items...")

    return translations

def main():
    all_items = []

    for f in glob.glob('fk/*.pdf'):
        all_items.extend(parse_mercadona_pdf(f))

    for f in glob.glob('fk/*.ods') + glob.glob('fk/*.xlsx') + glob.glob('fk/*.xls'):
        all_items.extend(parse_makro_excel(f))

    df = pd.DataFrame(all_items)
    print(f"Parsed {len(df)} total items.")

    def extract_month(date_str):
        if not date_str: return ''
        date_str = date_str.replace('-', '/')
        parts = date_str.split('/')
        if len(parts) >= 3:
            return f"{parts[2]}-{parts[1]}" if len(parts[2]) == 4 else f"20{parts[2][-2:]}-{parts[1]}"
        elif len(parts) == 2:
            return f"{parts[1]}-{parts[0]}"
        return date_str

    df['月份'] = df['日期'].apply(extract_month)

    # 1. Apply rules to map specific items to the requested '产品分类' (Product Category)
    df['产品分类 (手工规则)'] = df['原始名称'].apply(categorize_item)

    # 2. For items that didn't match our rules, we still translate the '原始名称' to '中文名称'
    # to serve as a generic product name
    uncategorized = df[df['产品分类 (手工规则)'].isnull()]['原始名称'].unique()
    translations = translate_descriptions(uncategorized)

    # Populate 中文名称 with the translation
    df['中文名称'] = df['原始名称'].map(translations)

    # Now set final 产品 (Product) column:
    # If it was matched by our rule, use the rule's result.
    # Otherwise, use the automated translation.
    df['产品 (统一分类)'] = df['产品分类 (手工规则)'].fillna(df['中文名称'])

    # Also populate 中文名称 for the manually categorized ones so it's not empty
    df['中文名称'] = df['中文名称'].fillna(df['产品 (统一分类)'])

    columns_order = [
        '商家', '月份', '日期', '发票号码', '产品 (统一分类)', '原始名称', '中文名称', '数量',
        '单价金额 (sin IVA)', '单品总计金额 (sin IVA)', 'IVA税率',
        '单价金额 (con IVA)', '单品总计金额 (con IVA)'
    ]
    df = df[columns_order]

    # Aggregate by Month, Merchant, and the Unified Product Category
    agg_df = df.groupby(['商家', '月份', '产品 (统一分类)']).agg({
        '数量': 'sum',
        '单品总计金额 (sin IVA)': 'sum',
        '单品总计金额 (con IVA)': 'sum'
    }).reset_index()

    agg_df = agg_df.rename(columns={
        '数量': '总用量 (数量)',
        '单品总计金额 (sin IVA)': '总计金额 (sin IVA)',
        '单品总计金额 (con IVA)': '总计金额 (con IVA)'
    })

    agg_df = agg_df.sort_values(by=['商家', '月份', '总计金额 (con IVA)'], ascending=[True, True, False])

    output_file = '订货统计.xlsx'

    with pd.ExcelWriter(output_file) as writer:
        df.to_excel(writer, sheet_name='详细列表 (Detailed List)', index=False)
        agg_df.to_excel(writer, sheet_name='用量与总计统计 (Usage & Totals)', index=False)

    print(f"Data successfully exported to {output_file} ({len(df)} detailed records, {len(agg_df)} aggregated records)")

if __name__ == "__main__":
    main()
