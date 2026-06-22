from flask import Flask, request, jsonify, send_file
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import json
import re
import io
import os

app = Flask(__name__)

HTML = open('index.html').read()

@app.route('/')
def index():
    return HTML

def parse_line_text(text):
    items = []
    current_team = 'UNKNOWN'
    for line in text.splitlines():
        line = line.strip()
        if not line:
            continue
        line_upper = line.upper()
        if re.search(r'ทีม\s*FO|^FO\s*:|^FO$|FO\s+สรุป|สรุป.*FO', line_upper):
            current_team = 'FO'
            continue
        if re.search(r'ทีม\s*MC|^MC\s*:|^MC$|MC\s+สรุป|สรุป.*MC', line_upper):
            current_team = 'MC'
            continue
        sku_match = re.search(r'(\d[\d\s]{5,14}\d)', line)
        if not sku_match:
            continue
        sku = re.sub(r'\s+', '', sku_match.group(1))
        rest = line[sku_match.end():]
        qty_match = re.search(r'[=\s\(]+(\d+)', rest)
        if not qty_match:
            continue
        qty = int(qty_match.group(1))
        team = current_team
        if re.search(r'\bFO\b', line_upper):
            team = 'FO'
        elif re.search(r'\bMC\b', line_upper):
            team = 'MC'
        items.append({'sku': sku, 'qty': qty, 'team': team})
    return items

@app.route('/api/parse-line', methods=['POST'])
def parse_line():
    text = request.json.get('text', '')
    try:
        items = parse_line_text(text)
        return jsonify({"ok": True, "items": items})
    except Exception as e:
        return jsonify({"ok": False, "items": [], "error": str(e)})

@app.route('/api/check-stock', methods=['POST'])
def check_stock():
    stock_file = request.files.get('stock')
    order_file = request.files.get('order')
    items_json = request.form.get('items', '[]')
    items = json.loads(items_json)

    if not stock_file or not items:
        return jsonify({"ok": False, "error": "ไม่มีไฟล์หรือ SKU"})

    df_stock = pd.read_excel(stock_file, header=1, dtype={'Seller SKU': str})
    df_order = pd.read_excel(order_file, dtype=str) if order_file else None

    order_sku_col = None
    if df_order is not None:
        for col in ['Fscode', 'รหัสสินค้า', 'SKU', 'sku']:
            if col in df_order.columns:
                order_sku_col = col
                break

    all_yokro_skus = {it['sku'] for it in items}
    wb = openpyxl.Workbook()
    border = Border(left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin'))
    hfill = PatternFill("solid", start_color="E91E63", end_color="E91E63")
    gfill = PatternFill("solid", start_color="C8E6C9", end_color="C8E6C9")
    yfill = PatternFill("solid", start_color="FFF9C4", end_color="FFF9C4")
    ofill = PatternFill("solid", start_color="FFE0B2", end_color="FFE0B2")
    rfill = PatternFill("solid", start_color="FFCDD2", end_color="FFCDD2")

    def hdr(ws, r, c, v):
        cell = ws.cell(row=r, column=c, value=v)
        cell.font = Font(bold=True, color="FFFFFF", name="Arial", size=10)
        cell.fill = hfill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border

    def dc(ws, r, c, v, fill, left=False, red=False, bold=False):
        cell = ws.cell(row=r, column=c, value=v)
        cell.font = Font(name="Arial", size=10, color="FF0000" if red else "000000", bold=bold)
        cell.fill = fill
        cell.alignment = Alignment(horizontal="left" if left else "center", vertical="center", wrap_text=True)
        cell.border = border

    ws1 = wb.active
    ws1.title = "เช็คสต็อกยกรอ"
    ws1.freeze_panes = 'A2'
    headers1 = ["ทีม","SKU","ชื่อสินค้า","MC","SH","Total","แจ้งยกรอ","รอแพ็คจริง","ผล","หมายเหตุ"]
    for ci,h in enumerate(headers1,1): hdr(ws1,1,ci,h)

    for ri, it in enumerate(items, 2):
        sku, rep, team = it['sku'], it['qty'], it['team']
        row = df_stock[df_stock['Seller SKU'] == sku]
        if row.empty:
            vals = [team,sku,"ไม่พบ SKU","-","-","-",rep,0,"❓ ไม่พบ","ไม่พบ SKU นี้ในระบบ"]
            fill = rfill
        else:
            r = row.iloc[0]
            mc,sh,total = int(r['MC']),int(r['SH']),int(r['Total'])
            ship = int(r['Shipment Unit']) if pd.notna(r['Shipment Unit']) else 0
            wait = int(r['Allowcate Unit']) - ship
            desc = str(r['Description'])[:45]
            if wait == 0:
                result,fill,note = "❌ ไม่มีรอแพ็ค",rfill,"มีของในคลัง แต่ไม่มี Allocate ค้าง"
            elif wait == rep:
                result,fill,note = "✅ ตรง",gfill,"ตรงกับที่แจ้ง"
            elif wait > rep:
                result,fill,note = "⚠️ รอแพ็คมากกว่า",yfill,f"รอแพ็คจริง {wait} > แจ้ง {rep}"
            else:
                result,fill,note = "⚠️ รอแพ็คน้อยกว่า",ofill,f"รอแพ็คจริง {wait} < แจ้ง {rep}"
            vals = [team,sku,desc,mc or'-',sh or'-',total,rep,wait,result,note]
        for ci,val in enumerate(vals,1):
            dc(ws1,ri,ci,val,fill,left=(ci==3))

    for ci,w in enumerate([8,14,38,8,8,8,12,12,18,28],1):
        ws1.column_dimensions[get_column_letter(ci)].width = w

    if df_order is not None and order_sku_col:
        ws2 = wb.create_sheet("Detail ออเดอร์")
        ws2.freeze_panes = 'A2'
        pfill = PatternFill("solid", start_color="FCE4EC", end_color="FCE4EC")
        wfill = PatternFill("solid", start_color="FFFFFF", end_color="FFFFFF")
        ord_cols = list(df_order.columns) + ['หมายเหตุ']
        for ci,h in enumerate(ord_cols,1): hdr(ws2,1,ci,h)
        team_map = {it['sku']: it['team'] for it in items}
        for ri,(_, row) in enumerate(df_order.iterrows(),2):
            sku = str(row.get(order_sku_col,'')).strip()
            is_yokro = sku in all_yokro_skus
            team = team_map.get(sku,'')
            remark = f"ยกรอ ({team})" if is_yokro else ''
            base_fill = wfill if ri%2==0 else pfill
            for ci,col in enumerate(ord_cols,1):
                val = remark if col=='หมายเหตุ' else (row[col] if col in row.index and pd.notna(row[col]) else '')
                dc(ws2,ri,ci,val,base_fill,red=is_yokro,bold=(col=='หมายเหตุ' and is_yokro),
                   left=(col in ['Description','รายละเอียด','Campaign']))
        ws2.column_dimensions['A'].width = 16
        for ci in range(2, len(ord_cols)+1):
            ws2.column_dimensions[get_column_letter(ci)].width = 16
        if 'Description' in ord_cols:
            ci = ord_cols.index('Description')+1
            ws2.column_dimensions[get_column_letter(ci)].width = 38
        if 'รายละเอียด' in ord_cols:
            ci = ord_cols.index('รายละเอียด')+1
            ws2.column_dimensions[get_column_letter(ci)].width = 38

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(buf, as_attachment=True,
                     download_name="สรุปยกรอ.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if __name__ == '__main__':
    app.run(port=5050, debug=True)
