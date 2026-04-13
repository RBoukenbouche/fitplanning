from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Font
import datetime, json, io, os

app = Flask(__name__)
CORS(app)

BASE = os.path.dirname(__file__)

# ── MONTANT EN LETTRES ────────────────────────────────────────────────────────
ONES = ['','One','Two','Three','Four','Five','Six','Seven','Eight','Nine',
        'Ten','Eleven','Twelve','Thirteen','Fourteen','Fifteen','Sixteen',
        'Seventeen','Eighteen','Nineteen']
TENS = ['','','Twenty','Thirty','Forty','Fifty','Sixty','Seventy','Eighty','Ninety']

def _b1000(n):
    if n < 20: return ONES[n]
    elif n < 100: return TENS[n//10]+(' '+ONES[n%10] if ONES[n%10] else '')
    else: return ONES[n//100]+' Hundred'+(' '+_b1000(n%100) if n%100 else '')

def amount_to_words(amount, currency='EUR'):
    amount = round(amount, 2)
    main = int(amount); cents = round((amount - main) * 100)
    cur_name = 'Euros' if currency == 'EUR' else 'US Dollars'
    parts = []
    if main // 1_000_000_000: parts.append(_b1000(main//1_000_000_000)+' Billion')
    if (main%1_000_000_000)//1_000_000: parts.append(_b1000((main%1_000_000_000)//1_000_000)+' Million')
    if (main%1_000_000)//1_000: parts.append(_b1000((main%1_000_000)//1_000)+' Thousand')
    if main%1_000: parts.append(_b1000(main%1_000))
    result = ' '.join(parts) if parts else 'Zero'
    if cents: result += f' And {_b1000(cents)} Cents'
    return result + f' {cur_name} Only'

# ── SAFE SET ──────────────────────────────────────────────────────────────────
def s(ws, row, col, val):
    try: ws.cell(row=row, column=col).value = val
    except AttributeError: pass

def clr(ws, r1, r2, c1=1, c2=9):
    for r in range(r1, r2+1):
        for c in range(c1, c2+1): s(ws, r, c, None)

# ── FORMATAGE DATES ───────────────────────────────────────────────────────────
def fmt_short(date_str):
    """2026-04-23 → Apr 23"""
    try: return datetime.datetime.strptime(date_str, '%Y-%m-%d').strftime('%b %d')
    except: return date_str

def add_days(date_str, n):
    try:
        d = datetime.datetime.strptime(date_str, '%Y-%m-%d')
        return (d + datetime.timedelta(days=n)).strftime('%Y-%m-%d')
    except: return date_str

def fmt_period(arr, dep):
    """2026-04-23, 2026-05-01 → Apr 23 - May 01, 2026"""
    try:
        a = datetime.datetime.strptime(arr, '%Y-%m-%d')
        d = datetime.datetime.strptime(dep, '%Y-%m-%d')
        return f"{a.strftime('%b %d')} - {d.strftime('%b %d, %Y')}"
    except: return f"{arr} - {dep}"

# ── CONSTRUCTION LIGNES HÉBERGEMENT ──────────────────────────────────────────
def build_accom_lines(accom_data, arrival, num_days, num_cats):
    """
    Regroupe les nuits consécutives même hôtel/ville/type
    Format : Apr 23-25 : Marrakech - Riad Palais - Double Room - 2 units
    """
    lines = []
    for c in range(num_cats):
        entries = []
        for i in range(num_days):
            dd = accom_data.get(str(i), accom_data.get(i, {}))
            cd = dd.get(str(c), dd.get(c, {}))
            entries.append({
                'city':     cd.get('city', '').strip(),
                'hotel':    cd.get('hotel', '').strip(),
                'roomType': cd.get('roomType', '').strip(),
                'units':    int(float(cd.get('units', '0') or 0)),
                'day':      i
            })

        # Regrouper les nuits consécutives même hôtel
        groups = []
        cur = None
        for e in entries:
            key = (e['city'], e['hotel'], e['roomType'], e['units'])
            if not e['hotel'] and not e['city']:
                if cur: groups.append(cur); cur = None
                continue
            if cur and (cur['city'], cur['hotel'], cur['roomType'], cur['units']) == key:
                cur['end_day'] = e['day']
            else:
                if cur: groups.append(cur)
                cur = {**e, 'start_day': e['day'], 'end_day': e['day']}
        if cur: groups.append(cur)

        for g in groups:
            if not g['hotel'] and not g['city']: continue
            d_from = fmt_short(add_days(arrival, g['start_day']))
            d_to   = fmt_short(add_days(arrival, g['end_day'] + 1))
            parts  = [f"{d_from}-{d_to}"]
            if g['city']:     parts.append(g['city'])
            if g['hotel']:    parts.append(g['hotel'])
            if g['roomType']: parts.append(g['roomType'])
            u = g['units']
            if u > 0: parts.append(f"{u} unit{'s' if u > 1 else ''}")
            lines.append(' - '.join(parts))
    return lines

# ── CONSTRUCTION LIGNES MEALS ─────────────────────────────────────────────────
def build_meal_lines(meals_data):
    """
    Une ligne par repas, triée par date
    Format : Apr 23 : La Maison Arabe — Welcome Dinner
    """
    lines = []
    sorted_meals = sorted(meals_data, key=lambda x: x.get('date', ''))
    for m in sorted_meals:
        desc = m.get('desc', '').strip()
        date = m.get('date', '')
        if not desc: continue
        date_lbl = fmt_short(date) if date else ''
        lines.append(f"{date_lbl}: {desc}" if date_lbl else desc)
    return lines

# ── CONSTRUCTION LIGNES TRANSPORT ────────────────────────────────────────────
def build_trans_lines(trans_data, arrival, num_days):
    lines = []
    for i in range(num_days):
        row = trans_data.get(str(i), trans_data.get(i, []))
        desc = row[0].strip() if isinstance(row, list) and len(row) > 0 else ''
        if desc:
            lines.append(f"{fmt_short(add_days(arrival, i))}: {desc}")
    return lines

# ── CONSTRUCTION LIGNES ACTIVITÉS ────────────────────────────────────────────
def build_act_lines(act_data):
    lines = []
    if not act_data: return lines
    # Format dict {0: [desc, rate, qty]} venant de actBody
    if isinstance(act_data, dict):
        for i in sorted(act_data.keys(), key=lambda x: int(x) if str(x).isdigit() else 0):
            row = act_data[i]
            if isinstance(row, list) and len(row) > 0 and row[0]:
                lines.append(str(row[0]).strip())
            elif isinstance(row, dict):
                desc = row.get('desc', '').strip()
                if desc: lines.append(desc)
        return lines
    # Format liste [{date, desc}]
    if isinstance(act_data, list):
        for a in act_data:
            if not isinstance(a, dict): continue
            desc = a.get('desc', '').strip()
            date = a.get('date', '')
            if not desc: continue
            date_lbl = fmt_short(date) if date else ''
            lines.append(f"{date_lbl}: {desc}" if date_lbl else desc)
    return lines

# ── CONSTRUCTION LIGNES EXTRAS ────────────────────────────────────────────────
def build_extra_lines(extras_data):
    lines = []
    for e in extras_data:
        desc = e.get('desc', '').strip()
        if desc: lines.append(desc)
    return lines

# ── GÉNÉRATION FACTURE ────────────────────────────────────────────────────────
def generate_invoice(data):
    currency  = data.get('currency', 'USD')
    is_eur    = currency == 'EUR'
    pax       = int(data.get('pax', 1))
    total_mad = float(data.get('totalSell', 0))
    rate      = float(data.get('eurRate', 10.5))
    num_days  = int(data.get('numDays', 0))
    num_cats  = int(data.get('numCats', 1))
    arrival   = data.get('arr', '')
    departure = data.get('dep', '')
    total_f   = int(((total_mad / rate) + 9) // 10) * 10  # Arrondi dizaine superieure
    per_pax   = int(((total_f / pax) + 9) // 10) * 10 if pax > 0 else 0  # Arrondi dizaine superieure

    ref       = data.get('ref', '')
    client    = data.get('client', '')
    contact   = data.get('contact', '')
    pax_names = data.get('paxNames', '')
    agent     = data.get('agentDisplay', '')
    notes     = data.get('notes', '')
    today     = datetime.datetime.today()
    period    = fmt_period(arrival, departure)

    # Construire les lignes depuis les données brutes du devis
    accom_lines = build_accom_lines(data.get('accomData', {}), arrival, num_days, num_cats)
    meal_lines  = build_meal_lines(data.get('mealsData', []))
    trans_lines = build_trans_lines(data.get('transData', {}), arrival, num_days)
    act_lines   = build_act_lines(data.get('actData', []))
    extra_lines = build_extra_lines(data.get('extData', []))
    inc_lines   = data.get('incLines', [])
    if not inc_lines:
        inc_lines = [
            'Accommodation at hotels as shown above or Similar',
            'All meals as indicated',
            'A/C Private deluxe vehicle at disposal',
            'High qualified English Speaking guide throughout',
            'All visits and activities as per the Itinerary',
            'Entrance fees to the monuments',
            'Water in the car during the tour',
            'All Local Taxes',
        ]
    # total_f est le montant arrondi — words sera recalculé après insertion dans Excel
    words       = amount_to_words(total_f, currency)

    tpl_file = 'template_vim.xlsx' if is_eur else 'template_sweet.xlsx'
    tpl_path = os.path.join(BASE, tpl_file)
    wb = load_workbook(tpl_path)
    ws = wb.active

    if is_eur:
        # VIM — En-tête (lignes fixes 1-6 non touchées)
        s(ws, 9,  1, 'Bill To:');    s(ws, 10, 1, client);    s(ws, 11, 1, contact)
        s(ws, 9,  7, 'Date:');       s(ws, 9,  8, today)
        s(ws, 10, 7, 'Invoice N°:'); s(ws, 10, 8, ref)
        s(ws, 11, 7, 'Ref N°:');     s(ws, 11, 8, pax_names)
        s(ws, 12, 7, 'Ref:');        s(ws, 12, 8, period)

        # Package per pax (row 17) : Unit=pax, Days=num_days, Daily Rate=per_pax
        s(ws, 17, 1, f'Package per person — {period}')
        s(ws, 17, 5, pax)       # UNIT
        s(ws, 17, 6, num_days)  # DAYS
        s(ws, 17, 7, per_pax)   # DAILY RATE
        s(ws, 17, 8, '=G17*E17') # GROSS = Daily Rate × Unit

        clr(ws, 18, 48)
        row = 20

        if accom_lines:
            s(ws, row, 1, 'Accommodation:')
            ws.cell(row=row, column=1).font = Font(bold=True)
            row += 1
            for l in accom_lines: s(ws, row, 1, l); row += 1

        if meal_lines:
            s(ws, row, 1, 'Meals:')
            ws.cell(row=row, column=1).font = Font(bold=True)
            row += 1
            for l in meal_lines: s(ws, row, 1, l); row += 1

        if act_lines:
            s(ws, row, 1, 'Activities & Experiences:'); row += 1
            for l in act_lines: s(ws, row, 1, l); row += 1

        if extra_lines:
            s(ws, row, 1, 'Other Services & Extras:'); row += 1
            for l in extra_lines: s(ws, row, 1, l); row += 1

        if inc_lines:
            row += 1
            s(ws, row, 1, 'Including:')
            ws.cell(row=row, column=1).font = Font(bold=True)
            row += 1
            for l in inc_lines: s(ws, row, 1, f'- {l}'); row += 1

        if notes: row += 1; s(ws, row, 1, notes); row += 1

        tr = max(row + 2, 49)
        s(ws, tr,   5, f'Total Invoice in {currency}')
        s(ws, tr,   8, f'=SUM(H17:H{tr-1})')
        s(ws, tr+1, 7, f'Deposit for confirmation {currency}')
        s(ws, tr+1, 8, f'=0.3*H{tr}')
        s(ws, tr+2, 7, f'Balance to be paid {currency}')
        s(ws, tr+2, 8, f'=H{tr}-H{tr+1}')
        # Total en lettres = montant réel total facture
        total_final = total_f
        words_final = amount_to_words(total_final, currency)
        s(ws, tr+4, 1, words_final)

    else:
        # SWEET — En-tête
        s(ws, 7,  1, 'Bill To:');    s(ws, 8,  1, client);    s(ws, 9,  1, contact)
        s(ws, 7,  7, 'Date:');       s(ws, 7,  8, today)
        s(ws, 8,  7, 'Invoice N°:'); s(ws, 8,  8, ref)
        s(ws, 9,  7, 'Ref N°:');     s(ws, 9,  8, pax_names)
        s(ws, 10, 7, 'Ref:');        s(ws, 10, 8, period)

        # Package per pax (row 15)
        s(ws, 15, 1, f'Package per person — {period}')
        s(ws, 15, 2, pax)       # UNIT
        s(ws, 15, 3, num_days)  # DAYS
        s(ws, 15, 4, per_pax)   # DAILY RATE
        s(ws, 15, 5, 0)         # VAT 0
        s(ws, 15, 6, '=B15*D15')
        s(ws, 15, 7, '=F15*E15')
        s(ws, 15, 8, '=F15+G15')

        clr(ws, 16, 44)
        row = 17

        if accom_lines:
            s(ws, row, 1, 'Accommodation:'); row += 1
            for l in accom_lines: s(ws, row, 1, l); row += 1

        if meal_lines:
            s(ws, row, 1, 'Meals:'); row += 1
            for l in meal_lines: s(ws, row, 1, l); row += 1

        if act_lines:
            s(ws, row, 1, 'Activities & Experiences:'); row += 1
            for l in act_lines: s(ws, row, 1, l); row += 1

        if extra_lines:
            s(ws, row, 1, 'Other Services & Extras:'); row += 1
            for l in extra_lines: s(ws, row, 1, l); row += 1

        if inc_lines:
            row += 1
            s(ws, row, 1, 'Including:')
            ws.cell(row=row, column=1).font = Font(bold=True)
            row += 1
            for l in inc_lines: s(ws, row, 1, f'- {l}'); row += 1

        if notes: row += 1; s(ws, row, 1, notes); row += 1

        tr = max(row + 2, 45)
        s(ws, tr,   4, 'Subtotal Excluding VAT')
        s(ws, tr,   8, f'=SUM(H15:H{tr-1})')
        s(ws, tr+1, 6, 'Total Output VAT')
        s(ws, tr+1, 8, '=G15')
        s(ws, tr+2, 5, f'Total Invoice in {currency}')
        s(ws, tr+2, 8, f'=SUM(H{tr}:H{tr+1})')
        s(ws, tr+3, 7, f'Deposit for confirmation {currency}')
        s(ws, tr+3, 8, f'=0.3*H{tr+2}')
        s(ws, tr+4, 7, f'Balance to be paid {currency}')
        s(ws, tr+4, 8, f'=H{tr+2}-H{tr+3}')
        # Total en lettres = montant réel total facture
        total_final = total_f
        words_final = amount_to_words(total_final, currency)
        s(ws, tr+6, 1, words_final)
        # Logo Sweet Spot
        logo_path = os.path.join(BASE, 'sweet_logo.png')
        if os.path.exists(logo_path):
            try:
                img = XLImage(logo_path)
                img.width = 150; img.height = 60
                ws.add_image(img, 'G1')
            except: pass

    # Sauvegarder en mémoire
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf, words

# ── EXPORT QUOTATION EXCEL ────────────────────────────────────────────────────
def generate_quotation_excel(data, sheet_title=None):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = (sheet_title or data.get('ref', 'Quotation'))[:31]

    # Styles
    navy='1A2E4A'; gold='C9A84C'; lb='E8EFF8'; white='FFFFFF'
    thin=Border(left=Side(style='thin',color='D8E2EF'),right=Side(style='thin',color='D8E2EF'),
                top=Side(style='thin',color='D8E2EF'),bottom=Side(style='thin',color='D8E2EF'))

    def st(cell,style='',num=False):
        cell.border=thin
        if style=='hdr':
            cell.font=Font(bold=True,color=white,size=10)
            cell.fill=PatternFill('solid',fgColor=navy)
            cell.alignment=Alignment(horizontal='center',vertical='center')
        elif style=='sec':
            cell.font=Font(bold=True,color=white,size=10)
            cell.fill=PatternFill('solid',fgColor=navy)
            cell.alignment=Alignment(horizontal='left',vertical='center')
        elif style=='gld':
            cell.font=Font(bold=True,color=navy,size=10)
            cell.fill=PatternFill('solid',fgColor=gold)
            cell.alignment=Alignment(horizontal='right',vertical='center')
        elif style=='lbl':
            cell.font=Font(bold=True,color='6B7C93',size=9)
            cell.fill=PatternFill('solid',fgColor=lb)
        elif style=='bold':
            cell.font=Font(bold=True,size=10)
        if num: cell.number_format='#,##0.00'
        return cell

    def w(row,col,val,style='',num=False,int_num=False):
        cell=ws.cell(row=row,column=col,value=val)
        st(cell,style,num)
        if int_num: cell.number_format='#,##0'
        return cell

    # Données
    currency=data.get('currency','EUR'); rate=float(data.get('rate',10.5))
    pax=int(data.get('pax',1)); num_days=int(data.get('numDays',0))
    num_cats=int(data.get('numCats',1)); arrival=data.get('arr','')
    margin_accom=float(data.get('marginAccom',0))
    margin_trans=float(data.get('marginTrans',20))
    margin_extras=float(data.get('marginExtras',20))
    margin_act=float(data.get('marginAct',20))

    r=1
    # TITRE
    ws.merge_cells(f'A{r}:H{r}')
    c=ws.cell(row=r,column=1,value='VISIT MOROCCO TRAVEL & EVENTS — QUOTATION')
    c.font=Font(bold=True,color=white,size=13); c.fill=PatternFill('solid',fgColor=navy)
    c.alignment=Alignment(horizontal='center',vertical='center')
    ws.row_dimensions[r].height=28; r+=1

    # INFOS GÉNÉRALES — le taux est dans une cellule référençable
    info_labels=[('Réf. Devis',data.get('ref',''),'Version',data.get('version',''),'Statut',data.get('status',''),'Date',data.get('quoteDate','')),
                 ('Client',data.get('client',''),'Contact',data.get('contact',''),'Passagers',pax,'Devise',currency),
                 ('Arrivée',arrival,'Départ',data.get('dep',''),'Jours',num_days,'Taux',rate)]
    
    rate_cell_ref = None  # On va stocker la référence de la cellule taux
    
    for row_idx, row_data in enumerate(info_labels):
        for i,v in enumerate(row_data):
            c=ws.cell(row=r,column=i+1,value=v); c.border=thin
            if i%2==0:
                c.font=Font(bold=True,color='6B7C93',size=9)
                c.fill=PatternFill('solid',fgColor=lb)
            else:
                c.font=Font(size=10)
                # Stocker la référence de la cellule taux (3ème ligne, colonne 8)
                if row_idx==2 and i==7:
                    rate_cell_ref = f'{get_column_letter(i+1)}{r}'
                    c.number_format='#,##0.00'
        r+=1
    r+=1

    # ── HÉBERGEMENT ──────────────────────────────────────────────────────────
    ws.merge_cells(f'A{r}:H{r}')
    st(ws.cell(row=r,column=1,value=f'🏨 HEBERGEMENT  (Marge: {margin_accom}%)'), 'sec')
    ws.row_dimensions[r].height=20; r+=1
    for i,h in enumerate(['Jour','Date','Hotel','Type Chambre','Unites','Tarif Unit. (MAD)','Total Achat (MAD)','Total Vente (MAD)']):
        w(r,i+1,h,'hdr')
    r+=1

    accom_data=data.get('accomData',{})
    accom_start=r
    for i in range(num_days):
        for c in range(num_cats):
            rd=(accom_data.get(str(i)) or accom_data.get(i) or {})
            cd=(rd.get(str(c)) or rd.get(c) or {})
            hotel=cd.get('hotel',''); room=cd.get('roomType','')
            units=float(cd.get('units',0) or 0); rate_u=float(cd.get('rate',0) or 0)
            if not hotel: continue
            dl=fmt_short(add_days(arrival,i)) if arrival else f'Jour {i+1}'
            w(r,1,f'Jour {i+1}'); w(r,2,dl); w(r,3,hotel); w(r,4,room)
            # Unités sans décimales
            w(r,5,units,'',False,True)
            w(r,6,rate_u,'',True)
            # Total achat = unités × tarif
            w(r,7,f'=E{r}*F{r}','',True)
            # Total vente arrondi à la dizaine = CEILING(achat/(1-marge), 10)
            if margin_accom < 100:
                w(r,8,f'=CEILING(G{r}/(1-{margin_accom/100}),10)','',True)
            else:
                w(r,8,f'=G{r}','',True)
            r+=1
    accom_end=r-1

    total_row_accom=r
    w(r,5,'TOTAL','bold')
    c7=ws.cell(row=r,column=7,value=f'=SUM(G{accom_start}:G{accom_end})')
    st(c7,'gld',True)
    c8=ws.cell(row=r,column=8,value=f'=SUM(H{accom_start}:H{accom_end})')
    st(c8,'gld',True)
    r+=2

    # ── GUIDE & TRANSPORT ────────────────────────────────────────────────────
    ws.merge_cells(f'A{r}:H{r}')
    st(ws.cell(row=r,column=1,value=f'🚗 GUIDE & TRANSPORT  (Marge: {margin_trans}%)'), 'sec')
    ws.row_dimensions[r].height=20; r+=1
    for i,h in enumerate(['Jour','Date','Description','Vehicule (MAD)','Guide (MAD)','Qte','Total Achat (MAD)','Total Vente (MAD)']):
        w(r,i+1,h,'hdr')
    r+=1
    trans_data=data.get('transData',{}); trans_start=r
    for i in range(num_days):
        rt=trans_data.get(str(i)) or trans_data.get(i) or []
        desc=rt[0] if rt else ''
        if not desc: continue
        veh=float(rt[1] or 0) if len(rt)>1 else 0
        guide=float(rt[2] or 0) if len(rt)>2 else 0
        qty=float(rt[3] or 1) if len(rt)>3 else 1
        dl=fmt_short(add_days(arrival,i)) if arrival else f'Jour {i+1}'
        w(r,1,f'Jour {i+1}'); w(r,2,dl); w(r,3,desc)
        w(r,4,veh,'',True); w(r,5,guide,'',True)
        w(r,6,qty,'',False,True)
        w(r,7,f'=(D{r}+E{r})*F{r}','',True)
        if margin_trans < 100:
            w(r,8,f'=CEILING(G{r}/(1-{margin_trans/100}),10)','',True)
        else:
            w(r,8,f'=G{r}','',True)
        r+=1
    trans_end=r-1
    trans_total_row=r
    w(r,5,'TOTAL','bold')
    c7=ws.cell(row=r,column=7,value=f'=SUM(G{trans_start}:G{trans_end})')
    st(c7,'gld',True)
    c8=ws.cell(row=r,column=8,value=f'=SUM(H{trans_start}:H{trans_end})')
    st(c8,'gld',True)
    r+=2

    # ── MEALS & SERVICES ─────────────────────────────────────────────────────
    ws.merge_cells(f'A{r}:H{r}')
    st(ws.cell(row=r,column=1,value=f'🍽️ MEALS & SERVICES  (Marge: {margin_extras}%)'), 'sec')
    ws.row_dimensions[r].height=20; r+=1
    for i,h in enumerate(['Description','Tarif Unit. (MAD)','Qte','Total Achat (MAD)','Total Vente (MAD)','','','']):
        w(r,i+1,h,'hdr' if i<5 else '')
    r+=1
    extras_data=data.get('extrasData',{}); extras_start=r
    for val in (extras_data.values() if isinstance(extras_data,dict) else []):
        desc=val[0] if len(val)>0 else ''
        tarif=float(val[1] or 0) if len(val)>1 else 0
        qty=float(val[2] or 0) if len(val)>2 else 0
        if not desc: continue
        w(r,1,desc); w(r,2,tarif,'',True)
        w(r,3,qty,'',False,True)
        w(r,4,f'=B{r}*C{r}','',True)
        if margin_extras < 100:
            w(r,5,f'=CEILING(D{r}/(1-{margin_extras/100}),10)','',True)
        else:
            w(r,5,f'=D{r}','',True)
        r+=1
    extras_end=r-1
    extras_total_row=r
    w(r,2,'TOTAL','bold')
    c4=ws.cell(row=r,column=4,value=f'=SUM(D{extras_start}:D{extras_end})')
    st(c4,'gld',True)
    c5=ws.cell(row=r,column=5,value=f'=SUM(E{extras_start}:E{extras_end})')
    st(c5,'gld',True)
    r+=2

    # ── ACTIVITIES ───────────────────────────────────────────────────────────
    ws.merge_cells(f'A{r}:H{r}')
    st(ws.cell(row=r,column=1,value=f'🎭 ACTIVITIES & EXPERIENCES  (Marge: {margin_act}%)'), 'sec')
    ws.row_dimensions[r].height=20; r+=1
    for i,h in enumerate(['Jour','Date','Description','Tarif Unit. (MAD)','Qte','Total Achat (MAD)','Total Vente (MAD)','']):
        w(r,i+1,h,'hdr' if i<7 else '')
    r+=1
    act_data=data.get('actData',{}); act_start=r
    for i in range(num_days):
        ra=act_data.get(str(i)) or act_data.get(i) or []
        desc=ra[0] if ra else ''
        tarif=float(ra[1] or 0) if len(ra)>1 else 0
        qty=float(ra[2] or 0) if len(ra)>2 else 0
        if not desc: continue
        dl=fmt_short(add_days(arrival,i)) if arrival else f'Jour {i+1}'
        w(r,1,f'Jour {i+1}'); w(r,2,dl); w(r,3,desc)
        w(r,4,tarif,'',True)
        w(r,5,qty,'',False,True)
        w(r,6,f'=D{r}*E{r}','',True)
        if margin_act < 100:
            w(r,7,f'=CEILING(F{r}/(1-{margin_act/100}),10)','',True)
        else:
            w(r,7,f'=F{r}','',True)
        r+=1
    act_end=r-1
    act_total_row=r
    w(r,4,'TOTAL','bold')
    c6=ws.cell(row=r,column=6,value=f'=SUM(F{act_start}:F{act_end})')
    st(c6,'gld',True)
    c7=ws.cell(row=r,column=7,value=f'=SUM(G{act_start}:G{act_end})')
    st(c7,'gld',True)
    r+=2

    # ── RÉCAP FINAL ──────────────────────────────────────────────────────────
    ws.merge_cells(f'A{r}:H{r}')
    st(ws.cell(row=r,column=1,value='💰 RECAP FINAL'), 'sec')
    ws.row_dimensions[r].height=20; r+=1
    for i,h in enumerate(['Section','Marge %','Total Achat (MAD)','Total Vente (MAD)',f'Total Vente ({currency})','','','']):
        w(r,i+1,h,'hdr' if i<5 else '')
    r+=1

    # Taux référencé depuis la cellule infos
    rate_ref = rate_cell_ref or str(rate)

    recap_start=r
    # Hébergement
    w(r,1,'Hebergement'); w(r,2,f'{margin_accom}%')
    w(r,3,f'=G{total_row_accom}','',True)
    w(r,4,f'=H{total_row_accom}','',True)
    w(r,5,f'=CEILING(D{r}/{rate_ref},10)','',True); r+=1
    # Transport
    w(r,1,'Transport'); w(r,2,f'{margin_trans}%')
    w(r,3,f'=G{trans_total_row}','',True)
    w(r,4,f'=H{trans_total_row}','',True)
    w(r,5,f'=CEILING(D{r}/{rate_ref},10)','',True); r+=1
    # Meals
    w(r,1,'Meals & Services'); w(r,2,f'{margin_extras}%')
    w(r,3,f'=D{extras_total_row}','',True)
    w(r,4,f'=E{extras_total_row}','',True)
    w(r,5,f'=CEILING(D{r}/{rate_ref},10)','',True); r+=1
    # Activities
    w(r,1,'Activities'); w(r,2,f'{margin_act}%')
    w(r,3,f'=F{act_total_row}','',True)
    w(r,4,f'=G{act_total_row}','',True)
    w(r,5,f'=CEILING(D{r}/{rate_ref},10)','',True); r+=1
    recap_end=r-1

    # Total général
    for j,v in enumerate(['TOTAL GENERAL','',
                           f'=SUM(C{recap_start}:C{recap_end})',
                           f'=SUM(D{recap_start}:D{recap_end})',
                           f'=CEILING(SUM(D{recap_start}:D{recap_end})/{rate_ref},10)',
                           '','','']):
        c2=ws.cell(row=r,column=j+1,value=v)
        st(c2,'gld')
        if j in[2,3,4]: c2.number_format='#,##0.00'
    total_gen_row=r; r+=1

    # Prix par personne
    for j,v in enumerate([f'PRIX PAR PERSONNE ({pax} pax)','','','',
                           f'=CEILING(E{total_gen_row}/{pax},10)',
                           '','','']):
        c2=ws.cell(row=r,column=j+1,value=v)
        st(c2,'gld' if j in[0,4] else '')
        if j==4: c2.number_format='#,##0.00'
    r+=1

    # Largeurs colonnes
    for i,width in enumerate([12,11,35,16,14,10,16,16]):
        ws.column_dimensions[get_column_letter(i+1)].width=width

    buf=io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf

# ── ROUTES ────────────────────────────────────────────────────────────────────
@app.route('/generate-invoice', methods=['POST'])
def route_invoice():
    try:
        data = request.get_json()
        buf, words = generate_invoice(data)
        currency = data.get('currency', 'USD')
        ref = data.get('ref', 'QT')
        client = data.get('client', 'Client').replace(' ', '_')
        filename = f"Invoice_{ref}_{client}.xlsx"
        return send_file(
            buf,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=filename
        )
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/export-quotation', methods=['POST'])
def route_export_quotation():
    try:
        data = request.get_json()
        buf = generate_quotation_excel(data)
        ref = data.get('ref', 'QT')
        client = data.get('client', 'Client').replace(' ', '_')
        filename = f"Quotation_{ref}_{client}.xlsx"
        return send_file(
            buf,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=filename
        )
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/export-all-versions', methods=['POST'])
def route_export_all_versions():
    try:
        from openpyxl import Workbook
        data = request.get_json()
        versions = data.get('versions', [])
        if not versions:
            return jsonify({'error': 'No versions provided'}), 400

        # Créer un workbook avec un onglet par version
        wb = None
        for i, version_data in enumerate(versions):
            buf = generate_quotation_excel(version_data)
            from openpyxl import load_workbook
            tmp_wb = load_workbook(buf)
            if wb is None:
                wb = tmp_wb
            else:
                # Copier la feuille dans le workbook principal
                ws_src = tmp_wb.active
                ws_new = wb.create_sheet(title=ws_src.title)
                for row in ws_src.iter_rows():
                    for cell in row:
                        new_cell = ws_new.cell(row=cell.row, column=cell.column, value=cell.value)
                        if cell.has_style:
                            new_cell.font = cell.font.copy()
                            new_cell.fill = cell.fill.copy()
                            new_cell.border = cell.border.copy()
                            new_cell.alignment = cell.alignment.copy()
                            new_cell.number_format = cell.number_format
                for col_dim in ws_src.column_dimensions.values():
                    ws_new.column_dimensions[col_dim.index].width = col_dim.width
                for row_dim in ws_src.row_dimensions.values():
                    ws_new.row_dimensions[row_dim.index].height = row_dim.height

        buf_out = io.BytesIO()
        wb.save(buf_out)
        buf_out.seek(0)

        ref = versions[0].get('ref', 'QT') if versions else 'QT'
        client = versions[0].get('client', 'Client').replace(' ', '_') if versions else 'Client'
        filename = f"AllVersions_{ref}_{client}.xlsx"
        return send_file(
            buf_out,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=filename
        )
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok', 'message': 'Visit Morocco Invoice Server running'})

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5055, debug=False)
