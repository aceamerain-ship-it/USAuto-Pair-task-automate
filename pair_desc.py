"""
Pair Description Processor — Python port of EU-Pair_Description_Processor.html
Processes Driver + Passenger HTML descriptions and produces merged output + attributes.
"""

import re

# ── Constants (from the HTML tool) ────────────────────────────────────────────
INFO_FINISH = ["Textured","Primed","Chrome","Satin","Bright","Glossy","Gloss","Matte",
               "Dull","Oil Painted","Oil-Painted","Paint to Match","Painted","Plain",
               "Powder-coated","Powdercoated","Phosphate Coated","Shiny","Smoked",
               "Smooth","Electroplated","Zinc Plated","Zinc-Plated","Zinc Coated",
               "Zinc-Coated","Paintable"]
COMMON_COLORS = ["Red","Blue","Green","Yellow","Black","White","Gray","Silver","Orange",
                 "Purple","Brown","Beige","Maroon","Navy","Gold","Ivory","Turquoise",
                 "Teal","Lime","Pink","Cyan","Magenta","Charcoal","Copper","Burgundy",
                 "Pearl","Olive","Tan","Mint","Lavender","Peach","Bronze","Coral",
                 "Slate","Rust","Ochre","Aquamarine","Platinum","Emerald","Amber",
                 "Sapphire","Crimson","Natural"]
EXTRA_POSITION = ["Inner","Outer","Lower","Upper","Rearward","Frontward","Forward"]
KEEP_LOWER = {"and","or","of","on","in","at"}

PREMIUM_AFTERMARKET = ("<br><br><b>Premium Aftermarket Replacement Parts:</b><br>"
    "- Every component undergoes a thorough evaluation and meticulous inspection to "
    "establish adherence to our strict safety standards and high-quality benchmarks."
    "<br><br><b>Engineered For Longevity:</b><br>"
    "- Crafted from top-tier materials, this product has undergone rigorous testing "
    "to provide maximum durability.<br><br>"
    "<b>Please confirm OEM # or Partslink # matches exactly before purchasing, "
    "otherwise, the item might not fit.")

# ── Helpers ────────────────────────────────────────────────────────────────────
def escape_regex(s):
    return re.escape(s)

def title_case_position(s):
    if not s: return s
    tokens = re.split(r'(\s+|,)', s)
    out = []
    for tok in tokens:
        if tok == ',' or re.match(r'^\s+$', tok):
            out.append(tok)
        elif tok.lower() in KEEP_LOWER:
            out.append(tok.lower())
        else:
            out.append(tok[0].upper() + tok[1:] if tok else tok)
    return re.sub(r'\s+,', ',', ''.join(out))

def convert_units(s):
    s = re.sub(r'(\d+)\s*mm\.?(?!\w)', r'\1 millimeters', s, flags=re.IGNORECASE)
    s = re.sub(r'(\d+)\s*in\.?(?!\w)', r'\1 inches', s, flags=re.IGNORECASE)
    s = re.sub(r'(\d+)\s*inch\.?(?!\w)', r'\1 inches', s, flags=re.IGNORECASE)
    s = re.sub(r'(\d+)\s*lb\.?(?!\w)', r'\1 Pounds', s, flags=re.IGNORECASE)
    s = re.sub(r'(\d+)\s*"', r'\1 inches', s)
    s = re.sub(r"\b x \b", ' by ', s)
    return s

def normalize_product_headers(s):
    s = re.sub(r'<b>(PRODUCT INFO):?\s*</b>', '<b>Product Info: </b>', s, flags=re.IGNORECASE)
    s = re.sub(r'<b>(PRODUCT INTERCHANGE):?\s*</b>', '<b>Product Interchange: </b>', s, flags=re.IGNORECASE)
    return s

def is_standalone_side(line):
    return bool(re.match(
        r'^\s*-?\s*(Driver Side|Passenger Side|Driver or Passenger Side|Driver and Passenger Side)\s*$',
        line, re.IGNORECASE))

def normalize_attribute_line(line):
    line = re.sub(r'^(Lens\s*Color:\s*)(.+)$', lambda m: m.group(1) + re.sub(r'\s*lens\s*$','',m.group(2),flags=re.IGNORECASE).strip(), line, flags=re.IGNORECASE)
    line = re.sub(r'^(Housing\s*Color:\s*)(.+)$', lambda m: m.group(1) + re.sub(r'\s*housing\s*$','',m.group(2),flags=re.IGNORECASE).strip(), line, flags=re.IGNORECASE)
    return line

def clean_lens_value(v):
    if not v or v == 'N/A': return v or 'N/A'
    return re.sub(r'\s*lens\s*$', '', v, flags=re.IGNORECASE).strip()

def clean_housing_value(v):
    if not v or v == 'N/A': return v or 'N/A'
    return re.sub(r'\s*housing\s*$', '', v, flags=re.IGNORECASE).strip()

def pluralize_parts(line):
    t = line.strip().lstrip('- ').strip()
    if t.lower().startswith('part: '):
        part = t[6:].strip()
        if part.lower() == 'glass': return 'Part: Glasses'
        words = part.split()
        if words and not words[-1].endswith('s'):
            words[-1] += 's'
        return 'Part: ' + ' '.join(words)
    return line

def normalize_for_set(s):
    t = s.strip().lstrip('- ').strip()
    t = re.sub(r'</?b>', '', t, flags=re.IGNORECASE)
    t = re.sub(r'\s+', ' ', t)
    t = re.sub(r'^(capa\s*(certified|certification))$', 'CAPA Certified', t, flags=re.IGNORECASE)
    t = re.sub(r'^with bulbs$', 'With Bulbs', t, flags=re.IGNORECASE)
    t = re.sub(r'^parts?link numbers?:', 'Partslink Numbers:', t, flags=re.IGNORECASE)
    t = re.sub(r'^oem numbers?:', 'OEM Numbers:', t, flags=re.IGNORECASE)
    if is_standalone_side(t): return 'skip-side-line'
    return t.lower()

# ── Position logic ─────────────────────────────────────────────────────────────
def parse_position_details(line):
    txt = re.sub(r'^-\s*', '', str(line or '').strip())
    txt = re.sub(r'^Position:\s*', '', txt, flags=re.IGNORECASE).strip()
    mains = []
    if re.search(r'\bFront\b', txt, re.IGNORECASE): mains.append('front')
    if re.search(r'\bRear\b', txt, re.IGNORECASE): mains.append('rear')
    side = ''
    if re.search(r'\bdriver\s*and\s*passenger\s*side\b', txt, re.IGNORECASE): side = 'driver and passenger side'
    elif re.search(r'\bdriver\s*or\s*passenger\s*side\b', txt, re.IGNORECASE): side = 'driver or passenger side'
    elif re.search(r'\bdriver\s*/\s*passenger\s*side\b', txt, re.IGNORECASE): side = 'driver or passenger side'
    elif re.search(r'\bdriver\s*side\b', txt, re.IGNORECASE): side = 'driver side'
    elif re.search(r'\bpassenger\s*side\b', txt, re.IGNORECASE): side = 'passenger side'
    extras = [e.lower() for e in EXTRA_POSITION if re.search(r'\b' + re.escape(e) + r'\b', txt, re.IGNORECASE)]
    return {'mains': mains, 'side': side, 'extras': extras}

def position_score(d):
    if not d: return 0
    return len(d.get('mains',[])) + (2 if d.get('side') else 0) + len(d.get('extras',[]))

def format_extra_position(extras):
    s = set(e.lower() for e in (extras or []))
    groups = []
    for a,b in [('inner','outer'),('lower','upper'),('rearward','frontward')]:
        if a in s and b in s: groups.append(f'{a} and {b}'); s.discard(a); s.discard(b)
        else:
            if a in s: groups.append(a); s.discard(a)
            if b in s: groups.append(b); s.discard(b)
    if 'forward' in s: groups.append('forward'); s.discard('forward')
    for e in EXTRA_POSITION:
        el = e.lower()
        if el in s: groups.append(el); s.discard(el)
    return ', '.join(groups)

def build_position(d):
    mains = d.get('mains', []) if d else []
    side  = (d.get('side') or '') if d else ''
    extras = d.get('extras', []) if d else []
    ms = set(m.lower() for m in mains)
    if 'front' in ms and 'rear' in ms: m_str = 'front and rear'
    elif 'front' in ms: m_str = 'front'
    elif 'rear' in ms: m_str = 'rear'
    else: m_str = ''
    ex_str = format_extra_position(extras)
    has_ex = bool(ex_str.strip())
    if m_str and not side and not has_ex: pos = m_str
    elif m_str and side and not has_ex: pos = f'{m_str}, {side}'
    elif not m_str and side and not has_ex: pos = side
    elif m_str and side and has_ex: pos = f'{m_str} {side}, {ex_str}'
    elif not m_str and side and has_ex: pos = f'{side}, {ex_str}'
    elif m_str and not side and has_ex: pos = f'{m_str}, {ex_str}'
    else: pos = ex_str
    return pos.strip().lower()

def merge_position(d1, d2):
    a = d1 or {'mains':[],'side':'','extras':[]}
    b = d2 or {'mains':[],'side':'','extras':[]}
    ms = set([x.lower() for x in (a.get('mains',[])+b.get('mains',[]))])
    mains = [m for m in ['front','rear'] if m in ms]
    ex_set = set([x.lower() for x in (a.get('extras',[])+b.get('extras',[]))])
    extras = [e.lower() for e in EXTRA_POSITION if e.lower() in ex_set]
    sA, sB = (a.get('side') or '').lower(), (b.get('side') or '').lower()
    if 'driver and passenger side' in (sA, sB): side = 'driver and passenger side'
    elif sA == 'driver or passenger side' and sB == 'driver or passenger side': side = 'driver and passenger side'
    elif {sA,sB} == {'driver side','passenger side'}: side = 'driver and passenger side'
    elif 'driver or passenger side' in (sA,sB) and (sA in ('driver side','passenger side') or sB in ('driver side','passenger side')): side = 'driver and passenger side'
    elif sA and sB and sA == sB: side = sA
    elif sA and not sB: side = sA
    elif sB and not sA: side = sB
    else: side = sA or sB or ''
    return {'mains': mains, 'side': side, 'extras': extras}

def handle_color_finish(info):
    parts = re.split(r'[;,/|]+', info)
    parts = [p.strip() for p in parts if p.strip()]
    finishes, colors, housing_colors = [], [], []
    for p in parts:
        fnd = None
        for f in INFO_FINISH:
            if re.search(r'\b' + re.escape(f) + r'\b', p, re.IGNORECASE):
                fnd = f; break
        remainder = re.sub(re.escape(fnd), '', p, flags=re.IGNORECASE).strip() if fnd else p
        if re.search(r'\bhousing\b', remainder, re.IGNORECASE):
            hc = re.sub(r'\bhousing\b', '', remainder, flags=re.IGNORECASE)
            hc = re.sub(r'\bcolor\b', '', hc, flags=re.IGNORECASE).strip()
            if hc: housing_colors.append(hc)
            if fnd and fnd not in finishes: finishes.append(fnd)
            continue
        if fnd:
            if fnd not in finishes: finishes.append(fnd)
            if remainder: colors.append(remainder)
        else:
            colors.append(remainder)
    result = {}
    if finishes: result['Finish'] = 'Finish: ' + ', '.join(finishes)
    if colors: result['Color'] = 'Color: ' + ', '.join(colors)
    if housing_colors: result['Housing Color'] = 'Housing Color: ' + ', '.join(housing_colors)
    return result

# ── Main processor ─────────────────────────────────────────────────────────────
def process_pair(input1_html, input2_html, partslink_merged, oem_merged, parts_name_merged=''):
    """
    Ports the mergeContent() function from EU-Pair_Description_Processor.html.
    Returns dict with all output column values.
    """
    # Decode XML entities back to HTML
    def decode(s):
        if not s: return ''
        return (s.replace('&lt;','<').replace('&gt;','>').replace('&amp;','&').replace('&quot;','"'))

    # Apply N/A rule: if value is 'N/A' or starts with prefix but has no real data, treat as empty
    def clean_input(v):
        if not v: return ''
        v = v.strip()
        if v.upper() == 'N/A': return ''
        # Values like "Partslink Numbers: N/A" or "OEM Numbers: N/A"
        cleaned = re.sub(r'^(Partslink Numbers:|OEM Numbers:)\s*N/A\s*$', '', v, flags=re.IGNORECASE).strip()
        return cleaned

    i1 = decode(input1_html or '')
    i2 = decode(input2_html or '')
    partslink = clean_input(partslink_merged or '')
    oem       = clean_input(oem_merged or '')
    parts_name = clean_input(parts_name_merged or '')

    i1 = convert_units(i1.strip().replace(' - ', ''))
    i2 = convert_units(i2.strip().replace(' - ', ''))
    i1 = normalize_product_headers(i1)
    i2 = normalize_product_headers(i2)

    # Extract Product Info blocks
    pi1_m = re.search(r'<b>Product Info: </b><br>(.+?)<br><br>', i1, re.IGNORECASE | re.DOTALL)
    pi2_m = re.search(r'<b>Product Info: </b><br>(.+?)<br><br>', i2, re.IGNORECASE | re.DOTALL)

    product_info_lines = []
    seen_keys = set()
    best_pos1 = None
    best_pos2 = None

    def process_pi_block(pi_match, best_pos):
        nonlocal product_info_lines, seen_keys
        if not pi_match: return best_pos
        for line in pi_match.group(1).split('<br>'):
            line = re.sub(r'Location:', 'Position:', line, flags=re.IGNORECASE)
            line = line.strip()
            if not line or 'Quantity Sold:' in line: continue
            cand = parse_position_details(line)
            if position_score(cand) > 0:
                if not best_pos or position_score(cand) > position_score(best_pos):
                    best_pos = cand
            if re.match(r'^Color\s*Finish\s*:', line, re.IGNORECASE):
                info = ':'.join(line.split(':')[1:]).strip()
                cf = handle_color_finish(info)
                for k in ('Finish','Color','Housing Color'):
                    if k in cf:
                        v = normalize_attribute_line(pluralize_parts(cf[k]))
                        key = normalize_for_set(v)
                        if key not in seen_keys:
                            seen_keys.add(key)
                            product_info_lines.append(v)
            else:
                v = normalize_attribute_line(pluralize_parts(line))
                key = normalize_for_set(v)
                if key not in seen_keys:
                    seen_keys.add(key)
                    product_info_lines.append(v)
        return best_pos

    best_pos1 = process_pi_block(pi1_m, best_pos1)
    best_pos2 = process_pi_block(pi2_m, best_pos2)

    # Merge positions
    has_pos = position_score(best_pos1) > 0 or position_score(best_pos2) > 0
    if has_pos:
        merged_pos = merge_position(best_pos1, best_pos2)
        merged_pos_str = build_position(merged_pos)
        product_info_lines = [l for l in product_info_lines
                              if not re.match(r'^Position:', l.strip().lstrip('- '), re.IGNORECASE)
                              and not is_standalone_side(l.strip().lstrip('- '))]
        if merged_pos_str:
            product_info_lines.append('Position: ' + merged_pos_str)

    # Sort: Brand → Part → Position → Other
    brand_lines, part_lines, pos_lines, other_lines = [], [], [], []
    for line in product_info_lines:
        t = line.strip().lstrip('- ')
        if re.match(r'^Brand:', t, re.IGNORECASE):
            if t not in brand_lines: brand_lines.append(t)
        elif re.match(r'^Part:', t, re.IGNORECASE):
            if t not in part_lines: part_lines.append(t)
        elif re.match(r'^Position:', t, re.IGNORECASE):
            # Keep best position
            parsed = build_position(parse_position_details(t))
            if not pos_lines: pos_lines.append(parsed)
        elif is_standalone_side(t):
            pass
        else:
            if t not in other_lines: other_lines.append(t)

    if not pos_lines:
        for line in product_info_lines:
            t = line.strip().lstrip('- ')
            if is_standalone_side(t) or re.search(r'Driver|Passenger|Front|Rear|Inner|Outer|Lower|Upper', t, re.IGNORECASE):
                p = build_position(parse_position_details(t))
                if p: pos_lines.append(p); break
        if not pos_lines: pos_lines.append('driver and passenger side')

    # CAPA normalization
    capa_found = False
    for i, l in enumerate(other_lines):
        t = l.strip().lstrip('- ').lower()
        if re.search(r'\bcapa\b', t):
            other_lines[i] = 'CAPA Certified'; capa_found = True
    if capa_found:
        first_capa = next(i for i,l in enumerate(other_lines) if l == 'CAPA Certified')
        other_lines = [l for i,l in enumerate(other_lines) if l != 'CAPA Certified' or i == first_capa]

    ordered = []
    seen_final = set()
    part_vals = set()

    def add(line):
        key = normalize_for_set(line)
        if key == 'skip-side-line': return
        if key in seen_final: return
        seen_final.add(key)
        ordered.append(normalize_attribute_line(line))

    for l in brand_lines: add(l)
    for l in part_lines:
        v = l.replace('Part: ','',1).strip().lower()
        # Skip if duplicate (singular/plural match)
        dup = any(v == p or v+'s' == p or (p.endswith('s') and v == p[:-1]) for p in part_vals)
        if not dup:
            part_vals.add(v)
            add(l)
    for p in pos_lines:
        pk = p.lower().strip()
        key = 'position:' + pk
        if key not in seen_final:
            seen_final.add(key)
            ordered.append('Position: ' + title_case_position(p))
    for l in other_lines:
        t = l.strip().lstrip('- ')
        if re.match(r'^Color\s*Finish:', t, re.IGNORECASE): continue
        if re.match(r'^(capa certified)$', t, re.IGNORECASE): add('CAPA Certified'); continue
        if re.match(r'^with bulbs$', t, re.IGNORECASE): add('With Bulbs'); continue
        add(t)

    ordered = [l for l in ordered if not re.match(r'^Sold as Pair$', l.strip(), re.IGNORECASE)]

    # Build HTML output
    html_out = ''
    plain_out = ''
    if ordered:
        html_out  += '<b>Product Info: </b><br>'
        plain_out += 'Product Info:\n'
        for line in ordered:
            clean = re.sub(r'^-\s*','', line.strip())
            clean = re.sub(r'(Part:\s*)-\s*', r'\1', clean)
            clean = re.sub(r'(Position:\s*)-\s*', r'\1', clean)
            html_out  += f'- {clean}<br>'
            plain_out += f'- {clean}\n'
        html_out  += '- Sold as Pair<br><br>'
        plain_out += '- Sold as Pair\n'

    html_out  += '<b>Product Interchange:</b><br>'
    plain_out += 'Product Interchange:\n'
    if partslink:
        html_out  += f'- {partslink}<br>'
        plain_out += f'- {partslink}\n'
    if oem:
        html_out  += f'- {oem}<br>'
        plain_out += f'- {oem}\n'
    if parts_name:
        html_out  += f'- Alternative Part Number: {parts_name}<br>'
        plain_out += f'- Alternative Part Number: {parts_name}\n'

    # ── Build table columns ────────────────────────────────────────────────────
    def gv(pattern, text, default='N/A'):
        m = re.search(pattern, text, re.IGNORECASE)
        return m.group(1).strip() if m else default

    # For attribute extraction, use a combined text without HTML tags
    combined = html_out.replace('<b>','').replace('</b>','').replace('<br>','\n')

    part_val     = gv(r'Part:\s*([^\n<]+)', combined)
    pos_val_raw  = gv(r'Position:\s*([^\n<]+)', combined)
    pos_val      = title_case_position(pos_val_raw) if pos_val_raw != 'N/A' else 'N/A'
    # Match Color: but not Lens Color: or Housing Color:
    attr4_m = re.search(r'(?:^|\n)- Color:\s*([^\n<]+)', combined, re.IGNORECASE)
    attr4_val = attr4_m.group(1).strip() if attr4_m else 'N/A'
    attr5_raw    = gv(r'Lens\s*Color:\s*([^\n<]+)', combined)
    attr6_val    = gv(r'Finish:\s*([^\n<]+)', combined)
    attr7_m = re.search(r'(?:^|\n)- Material:\s*([^\n<]+)', combined, re.IGNORECASE)
    attr7_val = attr7_m.group(1).strip() if attr7_m else 'N/A'
    attr8_val    = gv(r'Light\s*Source:\s*([^\n<]+)', combined)
    attr9_raw    = gv(r'Housing\s*Color:\s*([^\n<]+)', combined)
    attr5_val    = clean_lens_value(attr5_raw)
    attr9_val    = clean_housing_value(attr9_raw)
    has_capa     = bool(re.search(r'\bCAPA\b', combined, re.IGNORECASE))

    bp1 = f'{part_val}, {pos_val}'
    bp2_parts = []
    if partslink: bp2_parts.append(partslink)
    if oem: bp2_parts.append(oem)
    bp2 = ' / '.join(bp2_parts) if bp2_parts else 'N/A'

    bp3_parts = []
    if attr4_val != 'N/A': bp3_parts.append(attr4_val)
    if attr5_val != 'N/A': bp3_parts.append(attr5_val + ' Lens')
    if attr6_val != 'N/A': bp3_parts.append(attr6_val)
    if attr8_val != 'N/A': bp3_parts.append(attr8_val)
    if attr9_val != 'N/A': bp3_parts.append(attr9_val + ' Housing')
    bp3 = ', '.join(bp3_parts) if bp3_parts else 'Made From The Highest Quality Materials'

    bp4 = 'CAPA Certified' if has_capa else 'DOT & SAE Compliant'

    return {
        'pair_description':  html_out,          # V
        'premium_aftermarket': PREMIUM_AFTERMARKET, # W (already in SS — skip writing)
        'bullet_point_1':    bp1,               # Y
        'bullet_point_2':    bp2,               # Z
        'bullet_point_3':    bp3,               # AA
        'bullet_point_4':    bp4,               # AB
        'bullet_point_5':    '',                # AC (blank)
        'attr1_part':        part_val,          # AE
        'attr2_placement':   pos_val,           # AF
        'attr3_orientation': 'N/A',             # AG
        'attr4_color':       attr4_val,         # AH
        'attr5_lens_color':  attr5_val,         # AI
        'attr6_finish':      attr6_val,         # AJ
        'attr7_material':    attr7_val,         # AK
        'attr8_light_source':attr8_val,         # AL
        'attr9_housing_color':attr9_val,        # AM
        'attr10_cert':       'CAPA' if has_capa else 'N/A',  # AN
    }

