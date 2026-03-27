#!/usr/bin/env python3
import fitz, os, re, openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime

CN = {
    'MEDIA GALLINA':'半只鸡','JAMONCITOS DE POLLO':'小鸡腿','ALAS CONGELADAS':'冷冻鸡翅',
    'CUARTO TRASERO CONG':'冷冻鸡后腿','CONTRAMUSLO DESHUESA':'去骨鸡大腿',
    'CONTRAMUSLO SIN PIEL':'无皮鸡大腿','Contramuslos d':'去骨鸡大腿',
    'POLLO CERTIFICADO E':'优质整鸡','PATAS POLLO':'鸡爪','ALAS PARTIDAS':'鸡翅段',
    'ALAS BBQ CONG':'冷冻BBQ鸡翅','OREJA PRECOCINADA':'预制猪耳','CALLOS PRECOCINADOS':'预制牛肚',
    'CIGALA MEDIANA':'螯虾','GAMBÓN':'大虾','ATUN EN A. GIRASOL':'葵花油金枪鱼罐头',
    'Atún en aceite':'油浸金枪鱼','ESCAROLA ENTERA':'苦菊','ICEBERG':'冰山生菜',
    'GUISANTE FINO':'嫩豌豆','CHAMPIÑON BANDEJA P':'托盘蘑菇','APIO VERDE':'西芹',
    'JUDÍA PLANA 350 GR':'扁豆350g','JUDÍA PLANA 750 GR':'扁豆750g','COLIFLOR':'花椰菜',
    'AGUACATE BANDEJA':'托盘牛油果','GUACAMOLE 500 GR':'牛油果酱500g',
    'MANGO':'芒果','T.CHERRY 500 GR':'圣女果500g','FRAMBUESA':'树莓','MORA':'黑莓',
    'ARÁNDANO 225 GR':'蓝莓225g','NARANJA 5 KG.':'橙子5kg','NARANJA 5 KG':'橙子5kg',
    'BANANA':'香蕉','Banana':'香蕉','FRESÓN 1 KG.':'草莓1kg','FRESÓN 1 KG':'草莓1kg',
    'VINAGRE DE VINO BLAN':'白葡萄酒醋','TOMATE FRITO':'番茄酱','SALSA CESAR':'凯撒酱',
    'SALSA DE TRUFAS':'松露酱','Salsa de Trufa':'松露酱','Ketchup Hacend':'番茄酱(Hacendado)',
    'Salsa de Soja':'酱油','Queso lonchas':'片状奶酪','QUESO SANDWICH':'三明治奶酪',
    'Huevos de codo':'鹌鹑蛋','18 HUEVOS CODORNIZ':'鹌鹑蛋18颗','Leche condensa':'炼乳',
    'Nata montada a':'打发奶油','Agua mineral g':'矿泉水(大)','Refresco de na':'橙味汽水',
    'Refresco de li':'柠檬汽水','Refresco cola':'可乐汽水','Gaseosa Hacend':'苏打水',
    'Cerveza Clásic':'经典啤酒','LIMONADA LIGHT':'低糖柠檬水','Infusión fruto':'水果茶',
    'ZUMO NARANJA C/PULPA':'鲜榨橙汁(含果肉)','ZUMO DE NARANJA EXPR':'现榨橙汁',
    'Gofres clásico':'华夫饼','ACEITE DE OLIVA 0.4':'特级初榨橄榄油0.4%',
    'Papel higiénic':'卫生纸','Papel Multiuso':'多用纸巾','Pañuelos de pa':'手帕纸',
    'Lejía normal T':'漂白液','Detergente rop':'洗衣液','Ambientador sp':'空气清新剂',
    'Guantes de lát':'乳胶手套','DESENGRASANTE PISTOL':'去油污喷雾',
    'LOTE 3 BAYETAS MICRO':'微纤维抹布3件','PREPARACIÓN':'购物袋/服务费',
    'Preparación':'购物袋/服务费',
}

def cn(desc):
    if desc in CN: return CN[desc]
    du = desc.upper()
    for k,v in CN.items():
        if k.upper()==du: return v
    for k,v in CN.items():
        if du.startswith(k.upper()) or k.upper().startswith(du): return v
    if 'PARKING' in du: return '停车费'
    return ''

def pf(s):
    return float(s.strip().replace('.','').replace(',','.'))

def parse_pdf(path):
    doc = fitz.open(path)
    inv_num = inv_date = None
    all_lines = []
    for pg, page in enumerate(doc):
        lines = [l.strip() for l in page.get_text().split('\n')]
        if pg == 0:
            for l in lines:
                if 'Nº Factura:' in l and not inv_num: inv_num = l.split(':',1)[1].strip()
                if 'Fecha Factura:' in l and not inv_date: inv_date = l.split(':',1)[1].strip()
        start = 0
        for i,l in enumerate(lines):
            if 'Fecha Factura:' in l: start = i+1; break
        all_lines.extend(lines[start:])
    doc.close()

    SKIP = {'Descripción','Unid.','P.Unitario','B.Imp.','IVA','Cuota IVA','Importe',
            'FORMA DE PAGO TARJETA BANCARIA','FORMA DE PAGO','TARJETA BANCARIA',
            'Fecha factura simplificada:'}
    sections, cur_items = [], []
    cur_fnum = cur_fdate = None
    after_total = False
    i, n = 0, len(all_lines)

    while i < n:
        line = all_lines[i]
        if not line: i+=1; continue
        if line=='DETALLE (€)' or line.startswith('Total Factura'):
            if cur_items: sections.append({'fnum':cur_fnum,'fdate':cur_fdate,'items':cur_items})
            break
        if line in SKIP or line.startswith('PVP '): i+=1; continue
        if line=='TOTAL (€)':
            i+=1
            cnt=0
            while cnt<3 and i<n:
                v=all_lines[i]
                if not v: i+=1; continue
                try: pf(v); cnt+=1; i+=1
                except: break
            if cur_fnum:
                sections.append({'fnum':cur_fnum,'fdate':cur_fdate,'items':cur_items})
                cur_items=[]; cur_fnum=cur_fdate=None
            after_total=True; continue
        if line.startswith('Factura Simplificada:'):
            fnum = line.split(':',1)[1].strip()
            fdate = None
            for j in range(i+1, min(i+6,n)):
                if re.match(r'\d{2}/\d{2}/\d{4}', all_lines[j]):
                    fdate=all_lines[j]; break
            if after_total:
                sections.append({'fnum':fnum,'fdate':fdate,'items':cur_items})
                cur_items=[]; cur_fnum=cur_fdate=None; after_total=False
            else:
                cur_fnum=fnum; cur_fdate=fdate
            i+=3; continue
        if i+6<n:
            try:
                desc=line; us=all_lines[i+1]; pu=all_lines[i+2]
                bi=all_lines[i+3]; iv=all_lines[i+4]; cu=all_lines[i+5]; im=all_lines[i+6]
                units=int(us); p_unit=pf(pu); b_imp=pf(bi)
                if not (re.match(r'\d+%',iv) or iv=='(*)'): raise ValueError
                cuota=pf(cu); importe=pf(im)
                cur_items.append({'desc':desc,'units':units,'p_unit':p_unit,
                                  'b_imp':b_imp,'iva':iv,'cuota':cuota,'importe':importe})
                i+=7; after_total=False; continue
            except: pass
        i+=1
    return inv_num, inv_date, sections

# ─── Excel styles ──────────────────────────────────────────────────────────
def hdr_style(cell, bg='1F4E79', fg='FFFFFF', bold=True, sz=10):
    cell.font = Font(bold=bold, color=fg, size=sz)
    cell.fill = PatternFill('solid', fgColor=bg)
    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

def thin_border():
    s = Side(style='thin', color='AAAAAA')
    return Border(left=s, right=s, top=s, bottom=s)

def money_fmt(cell):
    cell.number_format = '#,##0.00'
    cell.alignment = Alignment(horizontal='right', vertical='center')

# Column headers
COLS = ['商家','发票编号','购物单据号','购物日期','月份',
        '原始名称','中文名称','数量',
        '单价 sin IVA','小计 sin IVA','IVA税率','IVA税额','合计 con IVA']

def write_detail_sheet(ws, rows):
    ws.row_dimensions[1].height = 28
    for ci, h in enumerate(COLS, 1):
        c = ws.cell(1, ci, h)
        hdr_style(c)
        c.border = thin_border()
    col_w = [8,22,20,12,8,28,20,6,14,14,8,12,14]
    for ci, w in enumerate(col_w, 1):
        ws.column_dimensions[get_column_letter(ci)].width = w

    alt = False
    for r, row in enumerate(rows, 2):
        bg = 'EBF3FB' if alt else 'FFFFFF'; alt = not alt
        for ci, val in enumerate(row, 1):
            c = ws.cell(r, ci, val)
            c.border = thin_border()
            c.alignment = Alignment(vertical='center', wrap_text=(ci in [6,7]))
            if ci in [9,10,12,13]:
                money_fmt(c)
            if ci == 11:
                c.alignment = Alignment(horizontal='center', vertical='center')
            if bg != 'FFFFFF':
                c.fill = PatternFill('solid', fgColor=bg)

def main():
    folder = '/home/user/er4wkdfrhui/fk'
    pdfs = sorted(f for f in os.listdir(folder) if f.endswith('.pdf'))

    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # remove default sheet

    all_rows = []

    for pdf_file in pdfs:
        path = os.path.join(folder, pdf_file)
        inv_num, inv_date, sections = parse_pdf(path)
        print(f'{pdf_file}: {inv_num}, {inv_date}, {len(sections)} sections')

        # Determine month label
        try:
            dt = datetime.strptime(inv_date, '%d/%m/%Y')
            month_label = dt.strftime('%Y-%m')
        except:
            month_label = '????'

        # Create per-invoice sheet
        sheet_name = month_label
        # Avoid duplicate sheet names
        existing = [s.title for s in wb.worksheets]
        if sheet_name in existing:
            sheet_name = sheet_name + '_2'
        ws = wb.create_sheet(sheet_name)

        # Invoice header
        ws.merge_cells('A1:M1')
        c = ws.cell(1,1, f'商家: MERCADONA S.A.    发票编号: {inv_num}    发票日期: {inv_date}')
        hdr_style(c, bg='1F4E79', sz=11)
        ws.row_dimensions[1].height = 22
        ws.merge_cells('A2:M2')
        ws.row_dimensions[2].height = 4

        ws.row_dimensions[3].height = 26
        for ci, h in enumerate(COLS, 1):
            c2 = ws.cell(3, ci, h)
            hdr_style(c2, bg='2E75B6')
            c2.border = thin_border()
        col_w = [8,22,20,12,8,28,20,6,14,14,8,12,14]
        for ci, w in enumerate(col_w, 1):
            ws.column_dimensions[get_column_letter(ci)].width = w

        row_idx = 4
        alt = False
        for sec in sections:
            fdate = sec['fdate'] or inv_date
            # Section sub-header
            ws.merge_cells(f'A{row_idx}:M{row_idx}')
            sc = ws.cell(row_idx, 1, f'购物单据: {sec["fnum"]}    购物日期: {fdate}')
            hdr_style(sc, bg='BDD7EE', fg='1F1F1F', bold=False, sz=9)
            sc.border = thin_border()
            ws.row_dimensions[row_idx].height = 16
            row_idx += 1

            sec_sin_iva = 0.0
            sec_cuota = 0.0
            sec_con_iva = 0.0

            for item in sec['items']:
                bg = 'F2F9FF' if alt else 'FFFFFF'; alt = not alt
                row_data = [
                    'MERCADONA', inv_num, sec['fnum'], fdate, month_label,
                    item['desc'], cn(item['desc']),
                    item['units'], item['p_unit'], item['b_imp'],
                    item['iva'], item['cuota'], item['importe']
                ]
                for ci, val in enumerate(row_data, 1):
                    cell = ws.cell(row_idx, ci, val)
                    cell.border = thin_border()
                    cell.alignment = Alignment(vertical='center', wrap_text=(ci in [6,7]))
                    if ci in [9,10,12,13]: money_fmt(cell)
                    if ci == 11: cell.alignment = Alignment(horizontal='center', vertical='center')
                    if bg != 'FFFFFF': cell.fill = PatternFill('solid', fgColor=bg)

                sec_sin_iva += item['b_imp']
                sec_cuota += item['cuota']
                sec_con_iva += item['importe']

                all_rows.append(row_data)
                row_idx += 1

            # Section total row
            for ci in range(1, 14):
                cell = ws.cell(row_idx, ci)
                cell.border = thin_border()
                cell.fill = PatternFill('solid', fgColor='FFF2CC')
            ws.cell(row_idx, 6, '小计')
            ws.cell(row_idx, 6).font = Font(bold=True)
            ws.cell(row_idx, 10, round(sec_sin_iva, 2)).number_format = '#,##0.00'
            ws.cell(row_idx, 10).font = Font(bold=True)
            ws.cell(row_idx, 12, round(sec_cuota, 2)).number_format = '#,##0.00'
            ws.cell(row_idx, 12).font = Font(bold=True)
            ws.cell(row_idx, 13, round(sec_con_iva, 2)).number_format = '#,##0.00'
            ws.cell(row_idx, 13).font = Font(bold=True)
            row_idx += 2  # blank line after each section

        # Freeze panes
        ws.freeze_panes = 'A4'

    # ─── All-data sheet ──────────────────────────────────────────────────
    ws_all = wb.create_sheet('全部明细', 0)
    write_detail_sheet(ws_all, all_rows)
    ws_all.freeze_panes = 'A2'

    # ─── Save ────────────────────────────────────────────────────────────
    out = '/home/user/er4wkdfrhui/fk/发票统计.xlsx'
    wb.save(out)
    print(f'\n✓ Saved: {out}')
    print(f'  Total rows: {len(all_rows)}')

if __name__ == '__main__':
    main()
