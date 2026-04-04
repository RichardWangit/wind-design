#!/usr/bin/env python3
"""Extract hub_v17.html into modular file structure."""
import os, re, html

SRC = '/home/user/wind-design/hub_v17.html'
with open(SRC, 'r', encoding='utf-8') as f:
    raw = f.read()
lines = raw.splitlines()

def mkdirs(*paths):
    for p in paths:
        os.makedirs(p, exist_ok=True)

mkdirs(
    '/home/user/wind-design/sections',
    '/home/user/wind-design/shared',
    '/home/user/wind-design/shared/db',
)

def write(path, content):
    with open(path, 'w', encoding='utf-8') as f:
        f.write(content)
    print(f'  wrote {path}  ({len(content)} chars)')

# ─────────────────────────────────────────────
# 1. Extract CSS (lines 8..239 = style block)
# ─────────────────────────────────────────────
style_start = raw.index('<style>\n') + len('<style>\n')
style_end   = raw.index('\n</style>', style_start)
css_content = raw[style_start:style_end]
write('/home/user/wind-design/shared/theme.css', css_content)

# ─────────────────────────────────────────────
# 2. Extract each iframe srcdoc  A–G
# ─────────────────────────────────────────────
IFRAMES = {
    'a': ('frame-a', '基本設計風速查詢'),
    'b': ('frame-b', '建築物用途係數查詢'),
    'c': ('frame-c', '地況種類查詢'),
    'd': ('frame-d', '地形係數Kzt計算'),
    'e': ('frame-e', '外風壓係數與內風壓係數查詢'),
    'f': ('frame-f', '設計風壓計算式查詢'),
    'g': ('frame-g', None),   # frame-g ends differently
}

for key, (fid, title) in IFRAMES.items():
    # find srcdoc start
    marker = f'<iframe id="{fid}" srcdoc="'
    # frame-g uses a different attribute order
    if key == 'g':
        marker_alt = f'<iframe id="{fid}" srcdoc="'
    start_pos = raw.find(marker)
    if start_pos == -1:
        # try without id first
        marker = f'id="{fid}"'
        idx = raw.find(marker)
        # find srcdoc after id
        start_pos = raw.find('srcdoc="', idx) + len('srcdoc="')
    else:
        start_pos = start_pos + len(marker)

    # find srcdoc end: either `" title=` or `" style=` closing the attribute
    # look for the pattern: newline + `"` at the beginning of the line followed by ` title=` or ` style=`
    end_marker_title = f'" title="{title}">' if title else None

    # Generic search: find `\n" ` patterns after start_pos that close the attribute
    # The closing is always `\n" title=` or `\n" style=`
    search_from = start_pos
    end_pos = -1
    for pattern in ['\n" title=', '\n" style=']:
        p = raw.find(pattern, search_from)
        if p != -1:
            if end_pos == -1 or p < end_pos:
                end_pos = p

    if end_pos == -1:
        print(f'  ERROR: could not find end of srcdoc for {fid}')
        continue

    encoded = raw[start_pos:end_pos]
    decoded = html.unescape(encoded)
    write(f'/home/user/wind-design/sections/sec-{key}.html', decoded)

# ─────────────────────────────────────────────
# 3. Extract section S inline body → sec-s.html
# ─────────────────────────────────────────────
sec_s_start_marker = '<div class="sec-body" id="body-s"'
sec_s_start = raw.find(sec_s_start_marker)
# body-s ends at </div>\n    </div>\n  </div> (end of content-panel-body)
# Find the matching close: look for the end of body-s div
# There are nested divs inside. Let's find the end by looking for </div>\n\n\n\n<script>
sec_s_end = raw.find('\n\n\n\n<script>', sec_s_start)
if sec_s_end == -1:
    sec_s_end = raw.find('\n<script>', sec_s_start)
# Get just the div content
sec_s_raw = raw[sec_s_start:sec_s_end].strip()

# sec-s.html is a standalone page that wraps the calc panel
sec_s_html = '''<!DOCTYPE html>
<html lang="zh-TW">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>設計風壓計算</title>
<link href="https://fonts.googleapis.com/css2?family=Noto+Serif+TC:wght@400;600;700;900&family=Noto+Sans+TC:wght@300;400;500&family=DM+Mono:wght@400;500&display=swap" rel="stylesheet">
<link rel="stylesheet" href="../shared/theme.css">
<style>
body { margin: 0; background: var(--paper); color: var(--ink); font-family: \'Noto Sans TC\', sans-serif; }
.sec-body { display: flex; flex: 1; flex-direction: column; min-height: 0; overflow: hidden; overflow-y: auto; background: var(--paper); }
</style>
</head>
<body>
'''

# Extract just the inner content (the calc-panel div)
calc_panel_start = sec_s_raw.find('<div id="calc-panel"')
calc_panel_content = sec_s_raw[calc_panel_start:]
# Remove the outer wrapping divs - just get the inner calc-panel
# Find the body-s wrapper close
# Actually let's include the script that's embedded in body-s (if any)
sec_s_html += '<div id="calc-panel-wrapper" style="flex:1;overflow-y:auto;background:var(--paper);">\n'
sec_s_html += calc_panel_content
sec_s_html += '\n</div>\n'

# Add sec-s script (buildParamGrid, runCalc will be moved here from hub.js)
# For now just add a bridge listener stub
sec_s_html += '''<script src="../shared/bridge.js"></script>
<script>
// Receive formula selection from hub
window.addEventListener('message', function(e) {
  if (!e.data || typeof e.data !== 'object') return;
  if (e.data.source === 'wind_pressure_formula') {
    window._formulaData = e.data;
    updateFormulaBanner(e.data);
  }
  if (e.data.source === 'hub_to_sec_s_results') {
    window._hubResults = e.data.results || {};
    buildParamGrid();
  }
});

function updateFormulaBanner(data) {
  const ref = document.getElementById('calc-formula-ref');
  const disp = document.getElementById('calc-formula-display');
  const note = document.getElementById('calc-formula-note');
  if (ref)  ref.textContent  = data.formula      || '—';
  if (disp) disp.textContent = data.formulaDisplay || '—';
  if (note) note.textContent = data.note          || '';
  buildParamGrid();
}

function getResultByType(type) {
  return (window._hubResults || {})[type] || null;
}
</script>
</body>
</html>
'''
write('/home/user/wind-design/sections/sec-s.html', sec_s_html)

# ─────────────────────────────────────────────
# 4. Extract main JS → hub.js
# ─────────────────────────────────────────────
script_start = raw.rfind('<script>\n') + len('<script>\n')
script_end   = raw.rfind('\n</script>')
hub_js = raw[script_start:script_end]
write('/home/user/wind-design/hub.js', hub_js)

# ─────────────────────────────────────────────
# 5. Build index.html skeleton
# ─────────────────────────────────────────────

# Extract the structural HTML between </style></head><body> and <script>
# We need:
#   - top-bar
#   - layout > sections-rail (with section headers A-G, S, R)
#   - layout > content-panel (with iframes using src=)
# Also the body-r inline div (results archive stays inline)

body_start = raw.find('<body>') + len('<body>')
body_end   = raw.rfind('<script>')

body_html = raw[body_start:body_end].strip()

# Replace srcdoc iframes with src= iframes
replacements = [
    ('a', '基本設計風速查詢'),
    ('b', '建築物用途係數查詢'),
    ('c', '地況種類查詢'),
    ('d', '地形係數Kzt計算'),
    ('e', '外風壓係數與內風壓係數查詢'),
    ('f', '設計風壓計算式查詢'),
    ('g', None),
]

# We'll do targeted replacements on the body_html
# Each block looks like:
# <div class="sec-body" id="body-X" ...>
#   <iframe id="frame-X" srcdoc="...long encoded content...
# " title="..."></iframe>
#     </div>
# Replace with:
# <div class="sec-body" id="body-X" ...>
#   <iframe id="frame-X" src="sections/sec-X.html" title="..."></iframe>
# </div>

import re as _re

for key, title in replacements:
    fid = f'frame-{key}'
    # Find the iframe tag start in body_html
    iframe_start = body_html.find(f'<iframe id="{fid}"')
    if iframe_start == -1:
        # Try different format for frame-g
        iframe_start = body_html.find(f'id="{fid}"')
        if iframe_start == -1:
            print(f'  WARNING: cannot find <iframe id="{fid}" in body_html')
            continue
        iframe_start = body_html.rfind('<iframe', 0, iframe_start)

    # Find where this iframe tag ends: look for ></iframe>
    iframe_end = body_html.find('></iframe>', iframe_start) + len('></iframe>')
    if iframe_end < len('></iframe>'):
        print(f'  WARNING: cannot find end of iframe {fid}')
        continue

    old_tag = body_html[iframe_start:iframe_end]

    # Build replacement
    title_attr = f' title="{title}"' if title else ''
    if key == 'g':
        # frame-g has different attributes
        new_tag = f'<iframe id="{fid}" src="sections/sec-{key}.html"{title_attr} style="width:100%;height:100%;border:none;background:#f5f0e8;display:block;"></iframe>'
    else:
        new_tag = f'<iframe id="{fid}" src="sections/sec-{key}.html"{title_attr}></iframe>'

    body_html = body_html[:iframe_start] + new_tag + body_html[iframe_end:]

# Also replace body-s inline content with an iframe
# Find the body-s div
body_s_start = body_html.find('<div class="sec-body" id="body-s"')
body_s_end_search = body_html.find('</div>\n\n    </div>\n  </div>', body_s_start)
if body_s_end_search == -1:
    # try alternate ending
    body_s_end_search = body_html.find('    </div>\n  </div>', body_s_start)

# Find the actual closing of body-s: count divs
if body_s_start != -1:
    pos = body_s_start
    depth = 0
    end_of_body_s = -1
    while pos < len(body_html):
        open_div = body_html.find('<div', pos)
        close_div = body_html.find('</div>', pos)
        if open_div == -1 and close_div == -1:
            break
        if open_div != -1 and (close_div == -1 or open_div < close_div):
            depth += 1
            pos = open_div + 4
        else:
            depth -= 1
            if depth == 0:
                end_of_body_s = close_div + len('</div>')
                break
            pos = close_div + 6

    if end_of_body_s != -1:
        old_body_s = body_html[body_s_start:end_of_body_s]
        # Replace the outer wrapper style but keep same id/class
        # Extract the style attr from the original
        style_match = _re.search(r'<div class="sec-body" id="body-s"([^>]*)>', old_body_s)
        style_attrs = style_match.group(1) if style_match else ''
        new_body_s = (f'<div class="sec-body" id="body-s"{style_attrs}>\n'
                      f'  <iframe id="frame-s" src="sections/sec-s.html" style="width:100%;height:100%;border:none;background:var(--paper);display:block;"></iframe>\n'
                      f'</div>')
        body_html = body_html[:body_s_start] + new_body_s + body_html[end_of_body_s:]

# Build the full index.html
index_html = '''<!DOCTYPE html>
<html lang="zh-TW">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>建築風力設計查詢整合系統</title>
<link href="https://fonts.googleapis.com/css2?family=Noto+Serif+TC:wght@400;600;700;900&family=Noto+Sans+TC:wght@300;400;500&family=DM+Mono:wght@400;500&display=swap" rel="stylesheet">
<link rel="stylesheet" href="shared/theme.css">
</head>
<body>
''' + body_html + '''

<script src="hub.js"></script>
</body>
</html>
'''
write('/home/user/wind-design/index.html', index_html)

print('\nDone.')
