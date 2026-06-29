"""
USAUTO Pair Project - Surgical XML Worker
Patches target sheets, evaluates all formulas, and stores real cached values.
Streams sharedStrings.xml to avoid loading 562MB into RAM.

Usage:
    python usauto_pair.py <input.xlsx> <output.xlsx>
"""

import sys, re, zipfile, os

# ── Pair Description Processor (inlined from pair_desc.py) ──────────────────
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

# ── Deduplication groups ────────────────────────────────────────────────────────
_DEDUP_GROUPS = [
    ('signal_light_no', [
        r'^Built In Signal Light:\s*(Without|No) Signal Light$',
        r'^Notes:\s*Without Signal Light$',
        r'^Without Signal Light$',
        r'^No Signal Light$',
    ]),
    ('blind_spot_no', [
        r'^Blind Spot Detection:\s*(Without|No) Blind Spot (Feature|Detection)$',
        r'^Without Blind Spot (Feature|Detection)$',
        r'^No Blind Spot (Feature|Detection)$',
    ]),
    ('memory_no', [
        r'^Memory Recall:\s*(Without|No) Memory$',
        r'^Without Memory$',
        r'^No Memory$',
    ]),
    ('puddle_no', [
        r'^Puddle Light Included:\s*(Without|No) Puddle Light$',
        r'^Without Puddle Light$',
        r'^No Puddle Light$',
    ]),
    ('auto_dim_no', [
        r'^Auto.?Dimming:\s*(Without|No) Auto.?Dimming$',
        r'^Auto Dimming:\s*(Without|No) Auto.?Dimm\w*$',
        r'^(Without|No) Auto.?Dimm\w*$',
        r'^Notes:\s*(Without|No) Auto.?Dimm\w*$',
    ]),
    ('glass_adj_power', [
        r'^Glass Adjustment Method:\s*Power( Adjust)?$',
        r'^Power Adjust$',
        r'^Power$',
    ]),
    ('color_black_base', [
        r'^Color:\s*Black Base$',
        r'^Notes:\s*Black [Bb]ase$',
        r'^Black [Bb]ase$',
    ]),
    ('heated_non', [
        r'^Heated:\s*Non-Heated$',
        r'^Non-Heated$',
    ]),
    ('pin_plug_3', [
        r'^GTN 3 Pin Plug$',
        r'^3 Pin Plug$',
    ]),
]

_COMPOUND_PUDDLE_AUTO = [
    r'^Without Puddle Light,\s*Without Auto-?Dimming$',
    r'^No Puddle Light,\s*(Auto Dimming:\s*)?No Auto-?Dimming$',
    r'^Puddle Light Included:\s*(Without|No) Puddle Light,\s*Auto Dimming:\s*(Without|No) Auto-?Dimming$',
    r'^Puddle Light Included:\s*(Without|No) Puddle Light,\s*Notes:\s*(Without|No) Auto-?Dimming$',
]

_SIDE_CAMERA_POSITIVE = r'^No Puddle Light,\s*No Auto-?Dimming,\s*Side View Camera$'
_SIDE_CAMERA_NEGATIVE = r'^Notes:\s*(Without|No) Auto.?Dimm\w+,\s*(Without|No) Side View Camera$'


def _match_any(line, patterns):
    for p in patterns:
        if re.match(p, line, re.IGNORECASE):
            return True
    return False


def _group_hit(line):
    if _match_any(line, _COMPOUND_PUDDLE_AUTO):
        return 'compound_puddle_auto'
    if re.match(_SIDE_CAMERA_POSITIVE, line, re.IGNORECASE):
        return 'side_camera_positive'
    if re.match(_SIDE_CAMERA_NEGATIVE, line, re.IGNORECASE):
        return 'side_camera_negative'
    for group_key, patterns in _DEDUP_GROUPS:
        if _match_any(line, patterns):
            return group_key
    return None


def _normalize_line(line):
    # "Heated: Heated" -> "Heated"
    if re.match(r'^Heated:\s*Heated$', line, re.IGNORECASE):
        return 'Heated'
    return line


def clean_pair_description(html):
    """Post-process pair description HTML to remove duplicate/redundant lines."""
    # Remove any residual #N/A lines (Excel error values that leaked into output)
    html = re.sub(r'(?:<br>- #N/A)+(<br>|$)', r'\1', html)
    html = re.sub(r'<br>$', '', html)
    if not html:
        return html

    parts = re.split(r'(<br>)', html)
    segments = []
    i = 0
    while i < len(parts):
        text = parts[i]
        br = parts[i+1] if (i+1 < len(parts) and parts[i+1] == '<br>') else ''
        segments.append([text, br])
        i += 2 if br else 1

    seen_groups = {}
    seen_puddle = False
    seen_auto   = False
    seen_compound_both = False
    seen_side_cam_pos  = False
    keep = [True] * len(segments)

    for idx, (text, br) in enumerate(segments):
        raw_t = text.strip()
        if not raw_t.startswith('- '):
            continue
        t = raw_t.lstrip('- ').strip()
        if not t:
            continue

        # Normalize
        norm = _normalize_line(t)
        if norm != t:
            segments[idx][0] = text[:text.index(t)] + norm + text[text.index(t)+len(t):]
            t = norm

        # Pre-check: Notes line containing auto-dim info (marks auto-dim as seen even if has extra info)
        if re.match(r'^Notes:', t, re.IGNORECASE) and re.search(r'(Without|No) Auto.?Dimm', t, re.IGNORECASE):
            if not re.match(r'^Notes:\s*(Without|No) Auto.?Dimm', t, re.IGNORECASE):
                # Has extra connector info — keep the line but mark auto-dim as seen
                seen_auto = True
        # Pre-check: standalone "Heated" dedup (after normalization both forms become "Heated")
        if t == 'Heated':
            if 'heated' in seen_groups:
                keep[idx] = False
                continue
            seen_groups['heated'] = idx

        group = _group_hit(t)
        if not group:
            continue

        if group == 'compound_puddle_auto':
            if seen_compound_both or (seen_puddle and seen_auto):
                keep[idx] = False
            else:
                seen_compound_both = seen_puddle = seen_auto = True

        elif group == 'side_camera_positive':
            seen_side_cam_pos = True

        elif group == 'side_camera_negative':
            if seen_side_cam_pos:
                keep[idx] = False

        elif group == 'puddle_no':
            if seen_puddle or seen_compound_both:
                keep[idx] = False
            else:
                seen_puddle = True

        elif group == 'auto_dim_no':
            # "Notes: X-connector, Without Auto-Dimming" has extra info -- keep but mark seen
            is_notes_with_extra = (
                re.match(r'^Notes:', t, re.IGNORECASE) and
                not re.match(r'^Notes:\s*(Without|No) Auto-?Dimm', t, re.IGNORECASE)
            )
            if is_notes_with_extra:
                seen_auto = True
                # If this Notes line also mentions auto-dimming at the end, mark as seen
                if re.search(r'(Without|No) Auto-?Dimm', t, re.IGNORECASE):
                    seen_auto = True  # already set, explicit for clarity
            elif seen_auto or seen_compound_both:
                keep[idx] = False
            else:
                seen_auto = True

        else:  # signal_light_no, blind_spot_no, memory_no, glass_adj_power
            if group in seen_groups:
                keep[idx] = False
            else:
                seen_groups[group] = idx

    result = ''
    for idx, (text, br) in enumerate(segments):
        if keep[idx]:
            result += text + br

    return result

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
def strip_usa(sku):
    """Remove leading USA- prefix from component SKU (e.g. USA-JMT35EL → JMT35EL)."""
    return re.sub(r'^USA-', '', (sku or '').strip(), flags=re.IGNORECASE)

def process_pair(input1_html, input2_html, partslink_merged, oem_merged, parts_name_merged='',
               comp_sku_driver='', comp_sku_passenger=''):
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
        if v.upper() in ('N/A', '#N/A', '#VALUE!', '#REF!', '#NAME?', '#DIV/0!'): return ''
        if v.upper() == 'N/A': return ''
        # Values like "Partslink Numbers: N/A" or "OEM Numbers: N/A"
        cleaned = re.sub(r'^(Partslink Numbers:|OEM Numbers:)\s*(?:N/A(?:,\s*)?)+$', '', v, flags=re.IGNORECASE).strip()
        return cleaned

    i1 = decode(input1_html or '')
    i2 = decode(input2_html or '')
    partslink = clean_input(partslink_merged or '')
    oem       = clean_input(oem_merged or '')
    parts_name = clean_input(parts_name_merged or '')

    # Component SKU fallback — used when both partslink and OEM are empty/N/A
    sku_d = strip_usa(comp_sku_driver)
    sku_p = strip_usa(comp_sku_passenger)
    sku_parts = [s for s in [sku_d, sku_p] if s]
    part_numbers_str = ', '.join(sku_parts) if sku_parts else ''
    use_sku_fallback = (not partslink and not oem) and bool(part_numbers_str)

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
    if use_sku_fallback:
        html_out  += f'- Part Numbers: {part_numbers_str}<br>'
        plain_out += f'- Part Numbers: {part_numbers_str}\n'
    else:
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
    if use_sku_fallback:
        bp2 = f'Part Numbers: {part_numbers_str}' if part_numbers_str else 'N/A'
    else:
        bp2_parts = []
        if partslink: bp2_parts.append(partslink)
        if oem: bp2_parts.append(oem)
        bp2 = ' / '.join(bp2_parts) if bp2_parts else 'N/A'

    # BP3: extract all detail lines from Product Info
    # (exclude Brand, Part, Position, Sold as Pair, CAPA, and pure negative/without lines)
    _bp3_exclude = re.compile(
        r'^(Brand:|Part:|Position:|Sold as Pair$|CAPA Certified$|DOT &|Without |No |Non-)',
        re.IGNORECASE)
    _pi_html = re.search(
        r'<b>Product Info: </b><br>(.*?)<br><br>',
        html_out, re.DOTALL)
    if _pi_html:
        _pi_lines = [l.lstrip('- ').strip()
                     for l in _pi_html.group(1).split('<br>')
                     if l.strip().startswith('- ')]
        _details = [l for l in _pi_lines
                    if l and not _bp3_exclude.match(l)
                    and l.lower() != 'sold as pair']
        bp3 = ', '.join(_details) if _details else 'Made From The Highest Quality Materials'
    else:
        # Fallback to attribute-based approach if no Product Info block found
        bp3_parts = []
        if attr4_val != 'N/A': bp3_parts.append(attr4_val)
        if attr5_val != 'N/A': bp3_parts.append(attr5_val + ' Lens')
        if attr6_val != 'N/A': bp3_parts.append(attr6_val)
        if attr8_val != 'N/A': bp3_parts.append(attr8_val)
        if attr9_val != 'N/A': bp3_parts.append(attr9_val + ' Housing')
        bp3 = ', '.join(bp3_parts) if bp3_parts else 'Made From The Highest Quality Materials'


    # If BP3 contains a 'Lens' detail, populate ATTR5 (Lens Color) from it
    def _extract_lens_from_bp3(bp3_str):
        if not bp3_str or bp3_str == 'Made From The Highest Quality Materials': return None
        for item in [i.strip() for i in bp3_str.split(',')]:
            m = re.match(r'^Lens\s*(?:Color)?:\s*(.+)$', item, re.IGNORECASE)
            if m: return re.sub(r'\s*lens\s*$', '', m.group(1), flags=re.IGNORECASE).strip() or None
            m2 = re.match(r'^(.+?)\s+Lens$', item, re.IGNORECASE)
            if m2: return m2.group(1).strip()
            if re.search(r'\bLens\b', item, re.IGNORECASE):
                val = re.sub(r'\s*lens\b', '', item, flags=re.IGNORECASE).strip()
                val = re.sub(r'^\s*(Color|Tint)?\s*:\s*', '', val, flags=re.IGNORECASE).strip()
                return val or None
        return None

    _lens_from_bp3 = _extract_lens_from_bp3(bp3)
    if _lens_from_bp3 and attr5_val == 'N/A':
        attr5_val = _lens_from_bp3


    bp4 = 'CAPA Certified' if has_capa else 'DOT & SAE Compliant'

    html_out = clean_pair_description(html_out)

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


# ── End of inlined pair_desc ─────────────────────────────────────────────────


# ── Argument handling ─────────────────────────────────────────────────────────
import glob

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

if len(sys.argv) >= 3:
    # Explicit paths passed (e.g. from .bat)
    INPUT_FILE  = sys.argv[1]
    OUTPUT_FILE = sys.argv[2]
else:
    # Auto-detect: look for .xlsx in input/ subfolder next to this script
    input_dir  = os.path.join(SCRIPT_DIR, 'input')
    output_dir = os.path.join(SCRIPT_DIR, 'output')

    if not os.path.isdir(input_dir):
        print(f"\n  ERROR: No input/ folder found at: {input_dir}")
        print(  "         Create an input/ folder and place your .xlsx file inside.")
        print(  "         Or run: python usauto_pair.py <input.xlsx> <output.xlsx>\n")
        sys.exit(1)

    xlsx_files = glob.glob(os.path.join(input_dir, '*.xlsx'))
    if not xlsx_files:
        print(f"\n  ERROR: No .xlsx file found in: {input_dir}\n")
        sys.exit(1)
    if len(xlsx_files) > 1:
        print(f"\n  WARNING: Multiple .xlsx files found in input/ — using first:")
        for f in xlsx_files: print(f"    {f}")
        print()

    INPUT_FILE  = xlsx_files[0]
    fname       = os.path.basename(INPUT_FILE)
    base, ext   = os.path.splitext(fname)
    OUTPUT_FILE = os.path.join(output_dir, base + '_output' + ext)

    if not os.path.isdir(output_dir):
        os.makedirs(output_dir)
        print(f"  Created output folder: {output_dir}")

if not os.path.isfile(INPUT_FILE):
    print(f"\n  ERROR: Input file not found:\n         {INPUT_FILE}\n"); sys.exit(1)
if not INPUT_FILE.lower().endswith('.xlsx'):
    print(f"\n  ERROR: Input file must be an .xlsx file:\n         {INPUT_FILE}\n"); sys.exit(1)
output_dir = os.path.dirname(os.path.abspath(OUTPUT_FILE))
if not os.path.isdir(output_dir):
    os.makedirs(output_dir)

# ── Constants ───────────────────────────────────────────────────────────────
# Description output columns for Pair Description Processor
PAIR_DESC_COLS = {
    'V':  ('2',  'pair_description'),
    'W':  ('2',  'premium_aftermarket'),  # write text value each time V is populated
    'Y':  ('2',  'bullet_point_1'),
    'Z':  ('2',  'bullet_point_2'),
    'AA': ('2',  'bullet_point_3'),
    'AB': ('2',  'bullet_point_4'),
    'AC': ('2',  'bullet_point_5'),
    'AE': ('2',  'attr1_part'),
    'AF': ('2',  'attr2_placement'),
    'AG': ('2',  'attr3_orientation'),
    'AH': ('2',  'attr4_color'),
    'AI': ('2',  'attr5_lens_color'),
    'AJ': ('2',  'attr6_finish'),
    'AK': ('2',  'attr7_material'),
    'AL': ('2',  'attr8_light_source'),
    'AM': ('2',  'attr9_housing_color'),
    'AN': ('2',  'attr10_cert'),
}
PAIR_DESC_TOUCH = set(PAIR_DESC_COLS.keys())
SHEET_IMG   = 'xl/worksheets/sheet3.xml'
SHEET_CA    = 'xl/worksheets/sheet2.xml'
SHEET_PHOTO = 'xl/worksheets/sheet4.xml'
SHEET_DESC  = 'xl/worksheets/sheet5.xml'
SHARED_STR  = 'xl/sharedStrings.xml'

COL_PASTE_VAL = 8    # col H
DATA_START    = 2
DATA_END      = 200  # overridden at runtime
SS_PASTE_HDR  = 1969891

# ── Helpers ───────────────────────────────────────────────────────────────────
PREFIX_RE = re.compile(r'^ITEMIMAGEURL\d+=\s*', re.IGNORECASE)
ROW_RE    = re.compile(r'<row[^>]+r="(\d+)"[^>]*>(.*?)</row>', re.DOTALL)
CELL_RE   = re.compile(r'<c r="([A-Z]+)(\d+)"([^>]*)>(.*?)</c>', re.DOTALL)
V_RE      = re.compile(r'<v>(.*?)</v>', re.DOTALL)

def col_letter(n):
    r = ''
    while n > 0:
        n, rem = divmod(n-1, 26); r = chr(65+rem) + r
    return r

def col_num(l):
    n = 0
    for ch in l: n = n*26 + (ord(ch)-64)
    return n

def strip_prefix(s):
    return PREFIX_RE.sub('', s.strip())

def xml_escape(s):
    return str(s).replace('&','&amp;').replace('<','&lt;').replace('>','&gt;').replace('"','&quot;')

def formula_escape(s):
    return str(s).replace('&','&amp;').replace('<','&lt;').replace('>','&gt;')

def clean_combined(raw):
    """Remove N/A entries, strip duplicates, normalize Combined Partslink or OEM string."""
    if not raw or raw == '#N/A': return None
    val_decoded = raw.replace('&#x2F;', '/').replace('&#47;', '/').replace('&amp;', '&')
    val_decoded = ' '.join(val_decoded.split())  # normalize all whitespace
    prefix = ''
    val = val_decoded
    for keyword in ('Partslink Numbers', 'OEM Numbers'):
        pat = re.match(r'^(' + re.escape(keyword) + r':\s*)(.*)', val_decoded, re.IGNORECASE)
        if pat:
            prefix = keyword + ': '  # normalize to exactly one space
            val = pat.group(2); break
    parts = [p.strip() for p in val.split(',')]
    parts = [p for p in parts if p and not re.match(r'^N/A$', p.strip(), re.IGNORECASE)]
    seen = set(); unique = []
    for p in parts:
        key = ' '.join(p.split()).upper()  # normalize for case-insensitive dedup
        if key not in seen:
            seen.add(key)
            unique.append(' '.join(p.split()))
    return (prefix + ', '.join(unique)) if unique else None

def h_style(rn):
    return '61' if rn==1 else '76' if rn==2 else '64'

def make_ss_cell(ref, ss_idx, style=None):
    s = f' s="{style}"' if style else ''
    return f'<c r="{ref}"{s} t="s"><v>{ss_idx}</v></c>'

def make_formula_cell(col, style, formula, rn, cell_type='str', cached='#N/A'):
    esc = formula_escape(formula)
    # cached values from SS are already XML-safe; plain text values are safe in <v>
    v   = cached
    return f'<c r="{col}{rn}" s="{style}" t="{cell_type}"><f>{esc}</f><v>{v}</v></c>'

def make_error_cell(col, style, formula, rn):
    """Formula cell that returns #N/A."""
    return make_formula_cell(col, style, formula, rn, cell_type='e', cached='#N/A')

def make_blank(ref, style):
    return f'<c r="{ref}" s="{style}"/>'

def remove_cells(cols_set, rn, body):
    col_refs = '|'.join(f'{c}{rn}' for c in cols_set)
    return re.sub(
        r'<c r="(?:' + col_refs + r')"(?:[^>]*/>[^<]*|[^>]*>.*?</c>)',
        '', body, flags=re.DOTALL)

def sort_row_cells(body):
    cells = [(col_num(m.group(1)), m.group(0))
             for m in re.finditer(r'<c r="([A-Z]+)\d+"(?:[^>]*/>[^<]*|[^>]*>.*?</c>)', body, re.DOTALL)]
    cells.sort(key=lambda x: x[0])
    return ''.join(c[1] for c in cells)

# ── Step 1: Read sheet XMLs ───────────────────────────────────────────────────
print("  [Python] Reading target sheets...")
with zipfile.ZipFile(INPUT_FILE, 'r') as z:
    sheet3_xml = z.read(SHEET_IMG).decode('utf-8')
    sheet4_xml = z.read(SHEET_PHOTO).decode('utf-8')
    sheet5_xml = z.read(SHEET_DESC).decode('utf-8')
    ca_xml     = z.read(SHEET_CA).decode('utf-8')

# ── Step 2: Parse sheet3 for SS indices ──────────────────────────────────────
print("  [Python] Parsing sheet3 for SKU / image URL indices...")
ss_indices_needed = set()
raw_rows = {}

for m in ROW_RE.finditer(sheet3_xml):
    rn = int(m.group(1))
    if rn < DATA_START: continue
    body = m.group(2)
    ss_idx = None; img_ss = None; img_val = None
    for cm in CELL_RE.finditer(body):
        col_l, attrs, cell_body = cm.group(1), cm.group(3), cm.group(4)
        v_m = V_RE.search(cell_body); val = v_m.group(1) if v_m else ''
        if col_l == 'A' and 't="s"' in attrs:
            ss_idx = int(val); ss_indices_needed.add(ss_idx)
        elif col_l == 'B':
            if 't="s"' in attrs:
                img_ss = int(val); ss_indices_needed.add(img_ss)
            else:
                img_val = val
    raw_rows[rn] = {'ss_idx': ss_idx, 'img_ss': img_ss, 'img_val': img_val or ''}

# Auto-detect DATA_END
DATA_END = max((rn for rn, rd in raw_rows.items() if rd['ss_idx'] is not None), default=200)
print(f"  [Python] Auto-detected DATA_END = {DATA_END}")

# ── Step 3: Parse CA Data Combined for lookup ─────────────────────────────────
print("  [Python] Reading CA Data Combined...")
# Columns needed: C (key), B, H, J, K, L, N, O, P, R, S, X, Z, AA, AB, AD, AE, AF, AH, AI
CA_TARGET = {'B','C','H','J','K','L','N','O','P','R','S','X','Z','AA','AB','AD','AE','AF','AH','AI'}
ca_raw = {}   # (rn, col) -> ('ss', idx) | ('val', str)

for m in ROW_RE.finditer(ca_xml):
    rn = int(m.group(1))
    if rn < 2: continue
    body = m.group(2)
    for col in CA_TARGET:
        pos = body.find(f'<c r="{col}{rn}"')
        if pos == -1: continue
        # Find the first > that closes the <c> opening tag
        tag_close = body.find('>', pos)
        if tag_close == -1: continue
        if body[tag_close-1] == '/': continue  # true self-closing <c .../> — skip
        end = body.find('</c>', tag_close)
        if end == -1: continue
        cell = body[pos:end+4]
        t = re.search(r't="([^"]+)"', cell)
        v = re.search(r'<v>(.*?)</v>', cell, re.DOTALL)
        if v:
            if t and t.group(1) == 's':
                idx = int(v.group(1)); ss_indices_needed.add(idx)
                ca_raw[(rn, col)] = ('ss', idx)
            else:
                ca_raw[(rn, col)] = ('val', v.group(1))

print(f"  [Python] CA Data cells: {len(ca_raw)}")

# ── Step 4: Stream sharedStrings.xml to resolve all needed indices ────────────
print(f"  [Python] Resolving {len(ss_indices_needed)} SS indices (streaming)...")
si_re = re.compile(r'<si>(.*?)</si>', re.DOTALL)
t_re  = re.compile(r'<t[^>]*>(.*?)</t>', re.DOTALL)
ss_map = {}
max_needed = max(ss_indices_needed) if ss_indices_needed else 0

# Read uniqueCount from header
with zipfile.ZipFile(INPUT_FILE, 'r') as z:
    with z.open(SHARED_STR) as f:
        hdr = f.read(500).decode('utf-8', errors='replace')
uc_m = re.search(r'uniqueCount="(\d+)"', hdr)
input_ss_count = int(uc_m.group(1)) if uc_m else 0
print(f"  [Python] Input SS uniqueCount: {input_ss_count}")

with zipfile.ZipFile(INPUT_FILE, 'r') as z:
    with z.open(SHARED_STR) as f:
        buf = b''; i = 0
        while len(ss_map) < len(ss_indices_needed):
            chunk = f.read(4 * 1024 * 1024)
            if not chunk: break
            buf += chunk
            last = 0
            for sm in si_re.finditer(buf.decode('utf-8', errors='replace')):
                if i in ss_indices_needed:
                    ss_map[i] = ''.join(t_re.findall(sm.group(1)))
                i += 1; last = sm.end()
                if i > max_needed: break
            buf = buf[last:]
            if i > max_needed: break

print(f"  [Python] Resolved {len(ss_map)} SS entries")

def resolve(entry):
    if entry is None: return None
    kind, val = entry
    return ss_map.get(val) if kind == 'ss' else val

# ── Step 5: Build CA Data lookup (keyed by CLX SKU Driver Side = col C) ──────
print("  [Python] Building CA Data lookup table...")
ca_lookup = {}   # clx_sku_driver -> {col: value}
for (rn, col), entry in ca_raw.items():
    key_entry = ca_raw.get((rn, 'C'))
    if not key_entry: continue
    key = resolve(key_entry)
    if not key: continue
    if key not in ca_lookup: ca_lookup[key] = {}
    val = resolve(entry)
    if val: ca_lookup[key][col] = val

print(f"  [Python] CA lookup keys: {len(ca_lookup)}")

# ── Step 6: Build rows_data ───────────────────────────────────────────────────
rows_data = {}
for rn in range(DATA_START, DATA_END + 1):
    rd = raw_rows.get(rn, {})
    ss_idx = rd.get('ss_idx')
    img_ss = rd.get('img_ss')
    img_val = rd.get('img_val', '')
    sku  = ss_map.get(ss_idx) if ss_idx is not None else None
    img  = ss_map.get(img_ss, img_val) if img_ss is not None else img_val
    tokens = [strip_prefix(p) for p in img.split(',') if p.strip()] if img else []
    rows_data[rn] = {'sku': sku, 'ss_idx': ss_idx, 'tokens': tokens}

# Build URL→SS index map
url_to_ss = {}; new_ss_urls = []
for rn in range(DATA_START, DATA_END + 1):
    for url in rows_data[rn]['tokens']:
        if url and url not in url_to_ss:
            url_to_ss[url] = input_ss_count + len(new_ss_urls)
            new_ss_urls.append(url)

data_count  = sum(1 for v in rows_data.values() if v['sku'])
max_tokens  = max((len(v['tokens']) for v in rows_data.values()), default=0)
print(f"  [Python] {data_count} rows with SKU | max {max_tokens} URLs/row")
print(f"  [Python] {len(new_ss_urls)} unique URLs → SS indices {input_ss_count}–{input_ss_count+len(new_ss_urls)-1}")

def get_row_url_ss(rn):
    return [url_to_ss[u] for u in rows_data[rn]['tokens'] if u in url_to_ss]

# ── Evaluation helpers ────────────────────────────────────────────────────────
def ca_val(sku, col):
    """Get CA Data Combined value for a given SKU and column. Returns None if not found."""
    return ca_lookup.get(sku, {}).get(col)

# Photowork formula → CA column mapping
PW_TO_CA = {
    'C': 'S',    # CLX SKU Passenger Side
    'D': 'B',    # Component SKU Driver (= Supplier Driver SKU in CA)
    'E': 'R',    # Component SKU Passenger (= CA!R)  — used internally
    'F': 'AH',   # Combined Partslink LH+RH
    'G': 'AI',   # Combined OEM LH+RH
    'H': 'J',    # Supplier Title Driver
    'I': 'Z',    # Supplier Title Passenger
}

def pw_eval(rn, pw_col):
    """Evaluate a Photowork formula column for a given row."""
    sku = rows_data[rn]['sku']
    if not sku: return None
    ca_col = PW_TO_CA.get(pw_col)
    if not ca_col: return None
    return ca_val(sku, ca_col)

def desc_eval(rn, desc_col):
    """Evaluate a Description formula column for a given row.
    Returns (value_str, is_error) where is_error=True means #N/A."""
    sku = rows_data[rn]['sku']
    if not sku: return (None, True)

    # Cols sourced from Photowork (which themselves pull from CA Data)
    PW_BACKED = {
        'C': ('pw', 'C'),   # Photowork!C = CA!S
        'D': ('pw', 'H'),   # Photowork!H = CA!J
        'E': ('pw', 'I'),   # Photowork!I = CA!Z
        'H': ('pw', 'D'),   # Photowork!D = CA!B
        'T': ('pw', 'F'),   # Photowork!F = CA!AH
        'U': ('pw', 'G'),   # Photowork!G = CA!AI
    }
    # Cols sourced directly from CA Data
    CA_BACKED = {
        'F': 'H',  'G': 'X',
        'I': 'R',  'J': 'K',  'K': 'L',
        'L': 'AA', 'M': 'AB',
        'N': 'N',  'O': 'O',  'P': 'P',
        'Q': 'AD', 'R': 'AE', 'S': 'AF',
    }

    if desc_col in PW_BACKED:
        src, ref = PW_BACKED[desc_col]
        val = pw_eval(rn, ref)
    elif desc_col in CA_BACKED:
        val = ca_val(sku, CA_BACKED[desc_col])
    else:
        return (None, True)

    if val is None: return (None, True)
    return (val, False)

# ── Step 7: Patch sheet3 ─────────────────────────────────────────────────────
print("  [Python] Patching sheet3 (Separate Image URLs Here)...")

def patch_sheet3(xml_str):
    def patch_row(m):
        rn, body, full = int(m.group(1)), m.group(2), m.group(0)
        if rn == 1:
            body_clean = re.sub(
                r'<c r="([A-Z]+)1"(?:[^>]*/>[^<]*|[^>]*>.*?</c>)',
                lambda cm: '' if col_num(cm.group(1)) >= COL_PASTE_VAL else cm.group(0),
                body, flags=re.DOTALL)
            nc = make_ss_cell('H1', SS_PASTE_HDR, style='61')
            patched = re.sub(r'spans="[^"]*"', 'spans="1:8"', full)
            return patched.replace(m.group(2), sort_row_cells(body_clean + nc))
        if DATA_START <= rn <= DATA_END:
            url_ss = get_row_url_ss(rn)
            body_clean = re.sub(
                r'<c r="([A-Z]+)' + str(rn) + r'"(?:[^>]*/>[^<]*|[^>]*>.*?</c>)',
                lambda cm: '' if col_num(cm.group(1)) >= COL_PASTE_VAL else cm.group(0),
                body, flags=re.DOTALL)
            if not url_ss:
                return full.replace(m.group(2), sort_row_cells(body_clean))
            nc = ''.join(make_ss_cell(col_letter(COL_PASTE_VAL+t)+str(rn), ss_idx,
                                       h_style(rn) if t==0 else None)
                         for t, ss_idx in enumerate(url_ss))
            max_c = COL_PASTE_VAL + len(url_ss) - 1
            patched = re.sub(r'spans="[^"]*"', f'spans="1:{max_c}"', full)
            return patched.replace(m.group(2), sort_row_cells(body_clean + nc))
        return full
    return ROW_RE.sub(patch_row, xml_str)

sheet3_patched = patch_sheet3(sheet3_xml)

# ── Step 8: Patch sheet4 (Photowork) ─────────────────────────────────────────
print("  [Python] Patching sheet4 (Photowork)...")

# Photowork formula columns with their CA Data source + styles
PHOTO_FORMULA_COLS = {
    'C':('79', "INDEX('CA Data Combined'!S:S,MATCH(Photowork!B{r},'CA Data Combined'!C:C,0))", 'S'),
    'D':('89', "INDEX('CA Data Combined'!B:B,MATCH(Photowork!B{r},'CA Data Combined'!C:C,0))", 'B'),
    'E':('79', "INDEX('CA Data Combined'!R:R,MATCH(Photowork!B{r},'CA Data Combined'!C:C,0))", 'R'),
    'F':('28', "INDEX('CA Data Combined'!AH:AH,MATCH(Photowork!B{r},'CA Data Combined'!C:C,0))", 'AH'),
    'G':('28', "INDEX('CA Data Combined'!AI:AI,MATCH(Photowork!B{r},'CA Data Combined'!C:C,0))", 'AI'),
    'H':('3',  "INDEX('CA Data Combined'!J:J,MATCH(Photowork!B{r},'CA Data Combined'!C:C,0))", 'J'),
    'I':('3',  "INDEX('CA Data Combined'!Z:Z,MATCH(Photowork!B{r},'CA Data Combined'!C:C,0))", 'Z'),
    'K':('15', 'CONCATENATE(D{r},"_master_SET_v2.JPG")', None),
    'M':('15', 'CONCATENATE(D{r},"_02_SET_v2.JPG")', None),
    'O':('15', 'CONCATENATE(D{r},"_03_SET_v2.JPG")', None),
    'Q':('15', 'CONCATENATE(D{r},"_04_SET_v2.JPG")', None),
    'S':('15', 'CONCATENATE(D{r},"_05_SET_v2.JPG")', None),
    'U':('15', 'CONCATENATE(D{r},"_06_SET_v2.JPG")', None),
    'W':('15', 'CONCATENATE(D{r},"_07_SET_v2.JPG")', None),
    'Y':('15', 'CONCATENATE(D{r},"_08_SET_v2.JPG")', None),
}

AA_STYLE_MAP = {
    **{r:'17' for r in [15,52,89,126,163,200]},
    **{r:'78' for r in [20,22,25,57,59,62,94,96,99,131,133,136,168,170,173]},
    **{r:'80' for r in [6,43,80,117,154,191]},
}
HLOOKUP_COLS  = ['AD','AE','AF','AG','AH','AI','AJ','AK','AL']
HLOOKUP_HDR   = {'AD':'_master','AE':'_01','AF':'_02','AG':'_03','AH':'_04','AI':'_05','AJ':'_06','AK':'_07','AL':'_08'}
PHOTO_TOUCH   = set(PHOTO_FORMULA_COLS) | {'B','J','AA','AB','AD','AE','AF','AG','AH','AI','AJ','AK','AL'}

def patch_photowork(xml_str):
    def patch_row(m):
        rn, body, full = int(m.group(1)), m.group(2), m.group(0)
        if rn < DATA_START or rn > DATA_END: return full
        body_clean = remove_cells(PHOTO_TOUCH, rn, body)
        sku = rows_data[rn]['sku']
        nc  = ''
        # B: CLX SKU Driver Side as SS ref
        if rows_data[rn]['ss_idx'] is not None:
            nc += f'<c r="B{rn}" s="98" t="s"><v>{rows_data[rn]["ss_idx"]}</v></c>'
        # INDEX/MATCH and CONCATENATE formula cols
        comp_sku = pw_eval(rn, 'D') if sku else None   # Component SKU Driver for CONCATENATE
        for col, (style, tmpl, ca_col) in PHOTO_FORMULA_COLS.items():
            formula = tmpl.replace('{r}', str(rn))
            if ca_col:  # INDEX/MATCH — evaluate
                raw_val = pw_eval(rn, col) if sku else None
                # For cols F and G (Combined Partslink/OEM), clean N/A and deduplicate
                if col in ('F', 'G'):
                    val = clean_combined(raw_val)
                    if val:
                        nc += make_formula_cell(col, style, formula, rn, cell_type='str', cached=val)
                    else:
                        # All N/A or no data — write "N/A" text (not #N/A error)
                        nc += make_formula_cell(col, style, formula, rn, cell_type='str', cached='N/A')
                else:
                    val = raw_val
                    if val:
                        nc += make_formula_cell(col, style, formula, rn, cell_type='str', cached=val)
                    else:
                        nc += make_error_cell(col, style, formula, rn)
            else:  # CONCATENATE — evaluate using Component SKU Driver (col D)
                if comp_sku:
                    suffix = re.search(r'"(_[^"]+)"', tmpl).group(1)
                    cached_val = comp_sku + suffix
                    nc += make_formula_cell(col, style, formula, rn, cell_type='str', cached=cached_val)
                else:
                    nc += make_error_cell(col, style, formula, rn)
        # J: blank
        nc += make_blank(f'J{rn}', '13')
        # AA, AB: blank
        nc += make_blank(f'AA{rn}', AA_STYLE_MAP.get(rn, '76'))
        nc += make_blank(f'AB{rn}', '80')
        # AD-AL: HLOOKUP with evaluated cached values
        tokens = rows_data[rn]['tokens']
        for col in ['AD'] + HLOOKUP_AE_AL:
            formula = f'HLOOKUP("*"&{col}$1&"*",\'Separate Image URLs Here\'!H{rn}:P{rn},1,FALSE)'
            header  = HLOOKUP_HDR[col]
            match_url = next((u for u in tokens if header.lower() in u.lower()), None)
            if col == 'AD' and rn == DATA_START and get_row_url_ss(rn):
                nc += f'<c r="AD{rn}" s="3" t="s"><v>{get_row_url_ss(rn)[0]}</v></c>'
            elif match_url:
                nc += make_formula_cell(col, '3', formula, rn, cell_type='str', cached=match_url)
            else:
                nc += make_error_cell(col, '3', formula, rn)
        return full.replace(m.group(2), sort_row_cells(body_clean + nc))
    return ROW_RE.sub(patch_row, xml_str)

HLOOKUP_AE_AL = ['AE','AF','AG','AH','AI','AJ','AK','AL']
sheet4_patched = patch_photowork(sheet4_xml)

# ── Step 8b: Run Pair Description Processor ─────────────────────────────────
print("  [Python] Running Pair Description Processor...")

def patch_pair_description(xml_str):
    """Run EU-Pair Description Processor logic for each row and write results to V, Y-AC, AE-AN."""
    def patch_row(m):
        rn, body, full = int(m.group(1)), m.group(2), m.group(0)
        if rn < DATA_START or rn > DATA_END: return full

        rd = rows_data.get(rn, {})
        if not rd.get('sku'): return full   # skip empty rows

        # Read evaluated F, G, T, U values from the already-patched xml
        def get_v(col):
            pos = body.find(f'<c r="{col}{rn}"')
            if pos == -1: return ''
            tag_close = body.find('>', pos)
            if tag_close == -1 or body[tag_close-1] == '/': return ''
            end = body.find('</c>', tag_close)
            if end == -1: return ''
            v = re.search(r'<v>(.*?)</v>', body[pos:end+4], re.DOTALL)
            return v.group(1) if v else ''

        f_val = get_v('F')   # DESCRIPTION Driver Side (XML-escaped HTML)
        g_val = get_v('G')   # DESCRIPTION Passenger Side
        t_val = get_v('T')   # Combined Partslink LH and RH
        u_val = get_v('U')   # Combined OEM LH and RH

        if not f_val and not g_val: return full

        h_val = get_v('H')   # Component SKU Driver Side
        i_val = get_v('I')   # Component SKU Passenger Side
        result = process_pair(f_val, g_val, t_val, u_val,
                              comp_sku_driver=h_val, comp_sku_passenger=i_val)

        # Remove existing pair desc cols
        body_clean = remove_cells(PAIR_DESC_TOUCH, rn, body)

        nc = ''
        # Cols that should always show a value (even N/A)
        always_value = {'AG','AH','AI','AJ','AK','AL','AM','AN'}
        for col, (style, key) in PAIR_DESC_COLS.items():
            val = result.get(key, '')
            if val and val not in ('',):
                esc = xml_escape(val)
                nc += f'<c r="{col}{rn}" s="{style}" t="str"><v>{esc}</v></c>'
            elif col in always_value:
                nc += f'<c r="{col}{rn}" s="{style}" t="str"><v>N/A</v></c>'
            else:
                nc += make_blank(f'{col}{rn}', style)

        return full.replace(m.group(2), sort_row_cells(body_clean + nc))

    return ROW_RE.sub(patch_row, xml_str)

# ── Step 9: Patch sheet5 (Description) ───────────────────────────────────────
print("  [Python] Patching sheet5 (Description)...")

DESC_FORMULA_COLS = {
    'C':('54','Photowork!C{r}'),
    'D':('35','Photowork!H{r}'),
    'E':('43','Photowork!I{r}'),
    'F':('21',"INDEX('CA Data Combined'!H:H,MATCH(Description!B{r},'CA Data Combined'!C:C,0))"),
    'G':('49',"INDEX('CA Data Combined'!X:X,MATCH(Description!B{r},'CA Data Combined'!C:C,0))"),
    'H':('20','Photowork!D{r}'),
    'I':('79',"INDEX('CA Data Combined'!R:R,MATCH(Description!B{r},'CA Data Combined'!C:C,0))"),
    'J':('20',"INDEX('CA Data Combined'!K:K,MATCH(Description!B{r},'CA Data Combined'!C:C,0))"),
    'K':('20',"INDEX('CA Data Combined'!L:L,MATCH(Description!B{r},'CA Data Combined'!C:C,0))"),
    'L':('50',"INDEX('CA Data Combined'!AA:AA,MATCH(Description!B{r},'CA Data Combined'!C:C,0))"),
    'M':('50',"INDEX('CA Data Combined'!AB:AB,MATCH(Description!B{r},'CA Data Combined'!C:C,0))"),
    'N':('20',"INDEX('CA Data Combined'!N:N,MATCH(Description!B{r},'CA Data Combined'!C:C,0))"),
    'O':('20',"INDEX('CA Data Combined'!O:O,MATCH(Description!B{r},'CA Data Combined'!C:C,0))"),
    'P':('20',"INDEX('CA Data Combined'!P:P,MATCH(Description!B{r},'CA Data Combined'!C:C,0))"),
    'Q':('50',"INDEX('CA Data Combined'!AD:AD,MATCH(Description!B{r},'CA Data Combined'!C:C,0))"),
    'R':('50',"INDEX('CA Data Combined'!AE:AE,MATCH(Description!B{r},'CA Data Combined'!C:C,0))"),
    'S':('50',"INDEX('CA Data Combined'!AF:AF,MATCH(Description!B{r},'CA Data Combined'!C:C,0))"),
    'T':('21','Photowork!F{r}'),
    'U':('59','Photowork!G{r}'),
}
DESC_TOUCH = set(DESC_FORMULA_COLS) | {'A','B'}

def patch_description(xml_str):
    def patch_row(m):
        rn, body, full = int(m.group(1)), m.group(2), m.group(0)
        if rn < DATA_START or rn > DATA_END: return full
        body_clean = remove_cells(DESC_TOUCH, rn, body)
        nc  = make_blank(f'A{rn}', '14')
        # B: CLX SKU Driver Side SS ref
        if rows_data[rn]['ss_idx'] is not None:
            nc += f'<c r="B{rn}" s="98" t="s"><v>{rows_data[rn]["ss_idx"]}</v></c>'
        # Formula cols with evaluated cached values
        for col, (style, tmpl) in DESC_FORMULA_COLS.items():
            formula = tmpl.replace('{r}', str(rn))
            val, is_err = desc_eval(rn, col)
            if is_err:
                if col in ('T', 'U'):
                    # Combined PL/OEM — write N/A text (not #N/A error) when no data
                    nc += make_formula_cell(col, style, formula, rn, cell_type='str', cached='N/A')
                else:
                    nc += make_error_cell(col, style, formula, rn)
            else:
                # For T and U (Combined Partslink/OEM via Photowork F/G), clean the value
                cleaned_val = clean_combined(val) if col in ('T', 'U') else val
                if cleaned_val:
                    nc += make_formula_cell(col, style, formula, rn, cell_type='str', cached=cleaned_val)
                elif col in ('T', 'U'):
                    # clean_combined stripped all N/A → write N/A text
                    nc += make_formula_cell(col, style, formula, rn, cell_type='str', cached='N/A')
                else:
                    nc += make_error_cell(col, style, formula, rn)
        return full.replace(m.group(2), sort_row_cells(body_clean + nc))
    return ROW_RE.sub(patch_row, xml_str)

sheet5_patched = patch_description(sheet5_xml)
sheet5_patched = patch_pair_description(sheet5_patched)

# ── Step 10: Write output ─────────────────────────────────────────────────────
print(f"  [Python] Writing → {OUTPUT_FILE}  (streaming SS, please wait)...")

new_total    = input_ss_count + len(new_ss_urls)
new_si_bytes = (''.join(f'<si><t>{xml_escape(u)}</t></si>' for u in new_ss_urls)).encode('utf-8')
CLOSE_TAG    = b'</sst>'
PATCHED_SMALL = {
    SHEET_IMG:   sheet3_patched.encode('utf-8'),
    SHEET_PHOTO: sheet4_patched.encode('utf-8'),
    SHEET_DESC:  sheet5_patched.encode('utf-8'),
}
TEMP = OUTPUT_FILE + '.tmp'

with zipfile.ZipFile(INPUT_FILE, 'r') as zin:
    with zipfile.ZipFile(TEMP, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            fn = item.filename
            if fn in PATCHED_SMALL:
                zout.writestr(item, PATCHED_SMALL[fn])
            elif fn == 'xl/calcChain.xml':
                pass
            elif fn == SHARED_STR:
                buf = b''; out_chunks = []; header_patched = False
                with zin.open(fn) as f:
                    while True:
                        chunk = f.read(4 * 1024 * 1024)
                        if not chunk: break
                        buf += chunk
                        if not header_patched:
                            h = buf[:500].decode('utf-8', errors='replace')
                            h = re.sub(r'uniqueCount="\d+"', f'uniqueCount="{new_total}"', h)
                            h = re.sub(r'count="\d+"', f'count="{new_total}"', h)
                            buf = h.encode('utf-8') + buf[500:]
                            header_patched = True
                        idx = buf.find(CLOSE_TAG)
                        if idx != -1:
                            out_chunks.append(buf[:idx])
                            out_chunks.append(new_si_bytes + CLOSE_TAG)
                            break
                        else:
                            keep = len(CLOSE_TAG) - 1
                            out_chunks.append(buf[:-keep] if len(buf) > keep else b'')
                            buf = buf[-keep:] if len(buf) > keep else buf
                zout.writestr(item, b''.join(out_chunks))
            else:
                zout.writestr(item, zin.read(fn))

os.replace(TEMP, OUTPUT_FILE)
print("  [Python] Done.")
