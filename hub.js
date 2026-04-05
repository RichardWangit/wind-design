const ALL = ['a','b','c','d','e','f','g','s','r'];
let active = null;

// ── 啟動：從 localStorage 還原上次狀態 ──────────────────────────────
HubState.restore();

// ── 依賴傳播訂閱（集中管理，取代各 addXxxResult 中的手動 postMessage）──
HubState.subscribe('wind', function(type, data) {
  if (!data) return;
  const gf = document.getElementById('frame-g');
  if (gf && gf.contentWindow) gf.contentWindow.postMessage({ source:'hub_to_gust_wind', speed: data.speed }, '*');
});

HubState.subscribe('terr', function(type, data) {
  if (!data) return;
  const gf = document.getElementById('frame-g');
  if (gf && gf.contentWindow) gf.contentWindow.postMessage({ source:'hub_to_gust_terrain', terrainId: data.terrainId }, '*');
  const df = document.getElementById('frame-d');
  if (df && df.contentWindow) df.contentWindow.postMessage({ source:'hub_to_kzt_exposure', terrainId: data.terrainId }, '*');
});

// 任何資料變動時同步更新計數並推送給 S 區
HubState.subscribe('*', function(type, data) {
  updateCount();
  if (type !== '*') pushResultsToSecS();
});

function toggleSection(which) {
  if (active === which) { closePanel(); } else { openPanel(which); }
}
function openPanel(which) {
  ALL.forEach(k => {
    const el = document.getElementById('sec-'+k);
    if (el) el.classList.toggle('expanded', k === which);
    const body = document.getElementById('body-'+k);
    if (body) body.style.display = (k === which) ? 'flex' : 'none';
  });
  active = which;
  ALL.forEach(k => setInd(k, k === which ? 'expanded' : 'collapsed'));
  const secEl = document.getElementById('sec-'+which);
  const badge = secEl && secEl.querySelector('.sec-badge');
  const title = secEl && secEl.querySelector('.sec-title');
  const cphBadge = document.getElementById('cph-badge');
  const cphTitle = document.getElementById('cph-title');
  if (cphBadge && badge) {
    cphBadge.textContent = badge.textContent.trim();
    cphBadge.style.background = window.getComputedStyle(badge).background;
  }
  if (cphTitle && title) cphTitle.textContent = title.textContent.trim();
  const accentMap = {a:'var(--accent)',b:'var(--accent2)',c:'var(--accent3)',
    d:'var(--accent5)',e:'var(--accent6)',f:'var(--accent7)',g:'#1a7a6a',
    s:'var(--accent8,#5a2d82)',r:'var(--gold)'};
  const cph = document.getElementById('cph');
  if (cph) cph.style.borderBottomColor = accentMap[which] || 'var(--accent)';
  document.getElementById('content-panel').classList.add('visible');
  if (which === 's') setTimeout(pushResultsToSecS, 300);
  if (which === 'g') {
    // Re-send cached A and C results to frame-g on panel open
    const gf = document.getElementById('frame-g');
    if (gf && gf.contentWindow) {
      const wa = HubState.get('wind');
      const wc = HubState.get('terr');
      if (wa) setTimeout(() => gf.contentWindow.postMessage({ source:'hub_to_gust_wind',    speed:     wa.speed     }, '*'), 300);
      if (wc) setTimeout(() => gf.contentWindow.postMessage({ source:'hub_to_gust_terrain', terrainId: wc.terrainId }, '*'), 350);
    }
  }
  if (which === 'd') {
    // Re-send cached C result to frame-d on panel open
    const df = document.getElementById('frame-d');
    if (df && df.contentWindow) {
      const wc = HubState.get('terr');
      if (wc) setTimeout(() => df.contentWindow.postMessage({ source:'hub_to_kzt_exposure', terrainId: wc.terrainId }, '*'), 300);
    }
  }
}
function closePanel() {
  if (!active) return;
  ALL.forEach(k => {
    const el = document.getElementById('sec-'+k);
    if (el) el.classList.remove('expanded');
    const body = document.getElementById('body-'+k);
    if (body) body.style.display = 'none';
  });
  active = null;
  ALL.forEach(k => setInd(k, 'collapsed'));
  document.getElementById('content-panel').classList.remove('visible');
}

function setInd(k, state) {
  const el = document.getElementById('ind-'+k);
  if (!el) return;
  el.textContent = state === 'expanded' ? '收束 ◀' : '展開 ▶';
}

// Auto-collapse query section → expand results after postMessage
function autoCollapse(which) { openPanel('r'); }

// ── RECEIVE postMessage ──
window.addEventListener('message', function(event) {
  const data = event.data;
  if (!data || typeof data !== 'object') return;
  if      (data.source === 'wind_speed_query')    { addWindResult(data);      if (active==='a') setTimeout(()=>autoCollapse('a'),700); }
  else if (data.source === 'occupancy_query')     { addOccResult(data);       if (active==='b') setTimeout(()=>autoCollapse('b'),700); }
  else if (data.source === 'terrain_query')       { addTerrainResult(data);   if (active==='c') setTimeout(()=>autoCollapse('c'),700); }
  else if (data.source === 'kzt_calculator')      { addKztResult(data);       if (active==='d') setTimeout(()=>autoCollapse('d'),700); }
  else if (data.source === 'wind_pressure_query') { addWindPressResult(data); if (active==='e') setTimeout(()=>autoCollapse('e'),700); }
  else if (data.source === 'wind_pressure_formula'){ addFormulaResult(data);  if (active==='f') setTimeout(()=>autoCollapse('f'),700); }
  else if (data.source === 'gust_effect_factor')  { addGustResult(data);      if (active==='g') setTimeout(()=>autoCollapse('g'),700); }
});

function addWindResult(data) {
  const id = 'r_'+Date.now();
  HubState.set('wind', data);
  hideEmpty();
  setPill('a', data.township+' · '+data.speed.toFixed(1)+' m/s');
  const card = document.createElement('div');
  card.className='result-card'; card.id=id;
  card.dataset.section='a';
  card.dataset.type='wind';
  card.innerHTML=`
    <div class="rc-section-badge s-a">A</div>
    <button class="rc-close" onclick="removeCard('${id}','wind')">✕</button>
    <div class="rc-type">🌀 基本設計風速</div>
    <div class="rc-loc">${data.island}<span class="slash">／</span>${data.county}<span class="slash">／</span><strong>${data.township}</strong></div>
    <div class="rc-value"><span class="rc-num">${data.speed.toFixed(1)}</span><span class="rc-unit">m/s</span></div>
    <div class="rc-zone">${data.zone}</div>
    <div class="rc-time">${formatTime(data.timestamp)}</div>`;
  replaceCard('a', card);
}

function addOccResult(data) {
  const id = 'r_'+Date.now();
  HubState.set('occ', data);
  hideEmpty();
  setPill('b', data.category+' · I = '+data.importanceFactor);
  const desc = data.itemDesc.length>32 ? data.itemDesc.substring(0,32)+'…' : data.itemDesc;
  const card = document.createElement('div');
  card.className='result-card type-occ'; card.id=id;
  card.dataset.section='b';
  card.dataset.type='occ';
  card.innerHTML=`
    <div class="rc-section-badge s-b">B</div>
    <button class="rc-close" onclick="removeCard('${id}','occ')">✕</button>
    <div class="rc-type">🏛 用途係數</div>
    <div class="rc-loc"><strong>${data.category}</strong><span class="slash">／</span>${desc}</div>
    <div class="rc-value"><span class="rc-unit" style="font-size:12px;color:var(--muted)">I =</span><span class="rc-num" style="color:var(--accent2)">${data.importanceFactor}</span></div>
    <div class="rc-zone">設計風速採 ${data.returnPeriod} 年回歸期</div>
    <div class="rc-time">${formatTime(data.timestamp)}</div>`;
  replaceCard('b', card);
}

function addTerrainResult(data) {
  const id = 'r_'+Date.now();
  HubState.set('terr', data);
  hideEmpty();
  setPill('c', '地況 '+data.terrainId+' · '+data.terrainName);
  // Show all 7 parameters
  const chips = data.parameters
    .map(p => `<span class="rc-param-chip">${p.symbol} = ${p.value}${p.unit!=='無因次' ? ' '+p.unit : ''}</span>`)
    .join('');
  const card = document.createElement('div');
  card.className='result-card type-terr'; card.id=id;
  card.style.width='100%';
  card.dataset.section='c';
  card.dataset.type='terr';
  card.innerHTML=`
    <div class="rc-section-badge s-c">C</div>
    <button class="rc-close" onclick="removeCard('${id}','terr')">✕</button>
    <div class="rc-type">🏔 地況種類</div>
    <div class="rc-loc"><strong>${data.terrainName}</strong></div>
    <div class="rc-value"><span class="rc-num" style="font-size:40px;color:var(--accent3)">${data.terrainId}</span></div>
    <div class="rc-zone" style="margin-top:2px">${data.description.length>50?data.description.substring(0,50)+'…':data.description}</div>
    <div class="rc-params">${chips}</div>
    <div class="rc-time">${formatTime(data.timestamp)}</div>`;
  replaceCard('c', card);
}

function addKztResult(data) {
  const id = 'r_'+Date.now();
  HubState.set('kzt', data);
  hideEmpty();
  setPill('d', data.terrain+' / 地況'+data.exposure+' · Kzt = '+data.results.Kzt);
  const r = data.results;
  const i = data.inputs;
  const card = document.createElement('div');
  card.className='result-card type-kzt'; card.id=id;
  card.style.width='100%';
  card.dataset.section='d';
  card.dataset.type='kzt';
  card.innerHTML=`
    <div class="rc-section-badge s-d">D</div>
    <button class="rc-close" onclick="removeCard('${id}','kzt')">✕</button>
    <div class="rc-type">⛰️ 地形係數 Kzt</div>
    <div class="rc-loc"><strong>${data.terrain}</strong><span class="slash">／</span>地況 ${data.exposure}</div>
    <div class="rc-value"><span class="rc-unit" style="font-size:12px;color:var(--muted)">Kzt =</span><span class="rc-num" style="color:var(--accent5)">${r.Kzt}</span></div>
    <div class="rc-params">
      <span class="rc-param-chip" style="background:rgba(26,107,138,0.1);color:var(--accent5)">K₁ = ${r.K1}</span>
      <span class="rc-param-chip" style="background:rgba(26,107,138,0.1);color:var(--accent5)">K₂ = ${r.K2}</span>
      <span class="rc-param-chip" style="background:rgba(26,107,138,0.1);color:var(--accent5)">K₃ = ${r.K3}</span>
    </div>
    <div class="rc-zone">H=${i.H}m　Lh=${i.Lh}m　x=${i.x}m　z=${i.z}m</div>
    <div class="rc-time">${formatTime(data.timestamp)}</div>`;
  replaceCard('d', card);
}

function addWindPressResult(data) {
  const id = 'r_'+Date.now();
  HubState.set('wind-press', data);
  hideEmpty();
  setPill('e', data.componentLabel+' · '+data.tableRef);

  // Build inputs rows
  const inps = data.inputs || [];
  const inpRows = inps.map(p =>
    `<tr>
      <td style="font-size:9px;padding:2px 5px;border-bottom:1px solid var(--border);color:var(--muted);white-space:nowrap">${p.label}</td>
      <td style="font-size:9px;padding:2px 5px;border-bottom:1px solid var(--border);font-family:'DM Mono',monospace;color:var(--ink)">${p.value}</td>
    </tr>`
  ).join('');

  // Build coefficient rows
  const coefs = data.coefficients || [];
  const coefRows = coefs.map(c => {
    const vStr = typeof c.value === 'number'
      ? (c.value >= 0 ? '+'+c.value.toFixed(3) : c.value.toFixed(3))
      : String(c.value);
    const isPos = typeof c.value === 'number' && c.value > 0;
    const isNeg = typeof c.value === 'number' && c.value < 0;
    const col = isPos ? 'color:#16a34a;font-weight:700'
               : isNeg ? 'color:#dc2626;font-weight:700'
               : 'color:#4f46e5;font-weight:700';
    return `<tr>
      <td style="font-size:9px;padding:3px 5px;border-bottom:1px solid var(--border);color:var(--muted)">${c.face}</td>
      <td style="font-size:9px;padding:3px 5px;border-bottom:1px solid var(--border);font-family:'DM Mono',monospace;color:var(--muted)">${c.coef}</td>
      <td style="font-size:11px;padding:3px 5px;border-bottom:1px solid var(--border);${col}">${vStr}</td>
    </tr>`;
  }).join('');

  const card = document.createElement('div');
  card.className='result-card type-wind'; card.id=id;
  card.style.width='100%';
  card.dataset.section='e';
  card.dataset.type='wind-press';
  card.innerHTML=`
    <div class="rc-section-badge s-e">E</div>
    <button class="rc-close" onclick="removeCard('${id}','wind-press')">✕</button>
    <div class="rc-type">🌬️ 風壓係數 · ${data.tableRef}</div>
    <div class="rc-loc"><strong>${data.buildingLabel}</strong></div>
    <div class="rc-loc" style="margin-top:3px;font-weight:400;font-size:10px;color:var(--muted)">${data.componentLabel}</div>
    ${inpRows.length > 0 ? `
    <div style="font-size:8px;letter-spacing:.1em;color:var(--muted);margin:8px 0 3px;font-family:'DM Mono',monospace">INPUT PARAMETERS</div>
    <table style="width:100%;border-collapse:collapse;margin-bottom:6px">
      ${inpRows}
    </table>` : ''}
    ${coefRows.length > 0 ? `
    <div style="font-size:8px;letter-spacing:.1em;color:var(--muted);margin:6px 0 3px;font-family:'DM Mono',monospace">COEFFICIENTS</div>
    <table style="width:100%;border-collapse:collapse">
      <tr>
        <th style="font-size:8px;padding:3px 5px;background:var(--border);text-align:left;letter-spacing:.06em">面</th>
        <th style="font-size:8px;padding:3px 5px;background:var(--border);text-align:left;letter-spacing:.06em">係數</th>
        <th style="font-size:8px;padding:3px 5px;background:var(--border);text-align:left;letter-spacing:.06em">數值</th>
      </tr>
      ${coefRows}
    </table>` : ''}
    <div class="rc-time">${formatTime(data.timestamp)}</div>`;
  replaceCard('e', card);
}

function addGustResult(data) {
  const id = 'r_'+Date.now();
  HubState.set('gust', data);
  hideEmpty();
  const isFlexible = data.buildingType === 'flexible';
  const symbol = isFlexible ? 'Gf' : 'G';
  setPill('g', symbol + ' = ' + data.valueStr);
  const card = document.createElement('div');
  card.className='result-card type-gust'; card.id=id;
  card.style.width='100%';
  card.dataset.section='g';
  card.dataset.type='gust';
  card.innerHTML=`
    <div class="rc-section-badge s-g">G</div>
    <button class="rc-close" onclick="removeCard('${id}','gust')">✕</button>
    <div class="rc-type">🌪️ 陣風反應因子 · §2.7</div>
    <div class="rc-loc"><strong>${data.label}</strong></div>
    <div class="rc-value">
      <span class="rc-unit" style="font-size:12px;color:var(--muted)">${symbol} =</span>
      <span class="rc-num" style="color:#1a7a6a">${data.valueStr}</span>
    </div>
    <div class="rc-zone">${data.desc}</div>
    <div class="rc-time">${formatTime(data.timestamp)}</div>`;
  replaceCard('g', card);
}

function addFormulaResult(data) {
  const id = 'r_'+Date.now();
  HubState.set('formula', data);
  hideEmpty();
  const mainFormulas  = data.formulas.filter(f => !f.isExtra);
  const extraFormulas = data.formulas.filter(f =>  f.isExtra);
  const pillText = data.subcategory + (mainFormulas.length ? ' · 公式'+mainFormulas.map(f=>f.code).join('/') : '');
  setPill('f', pillText);

  function buildFormulaRows(fList) {
    return fList.map(f => `
      <tr>
        <td style="font-size:9px;padding:3px 5px;border-bottom:1px solid var(--border);font-family:'DM Mono',monospace;color:var(--accent7);font-weight:700;white-space:nowrap">公式 ${f.code}</td>
        <td style="font-size:9px;padding:3px 5px;border-bottom:1px solid var(--border);color:var(--ink);font-weight:500">${f.name}</td>
      </tr>
      <tr>
        <td colspan="2" style="font-size:10px;padding:3px 5px 7px 5px;border-bottom:1px solid var(--border);font-family:'DM Mono',monospace;color:var(--accent7);background:rgba(45,80,22,0.05)">${f.display}</td>
      </tr>`).join('');
  }

  const mainRows  = buildFormulaRows(mainFormulas);
  const extraRows = buildFormulaRows(extraFormulas);

  const card = document.createElement('div');
  card.className='result-card type-formula'; card.id=id;
  card.style.width='100%';
  card.dataset.section='f';
  card.dataset.type='formula';
  card.innerHTML=`
    <div class="rc-section-badge s-f">F</div>
    <button class="rc-close" onclick="removeCard('${id}','formula')">✕</button>
    <div class="rc-type">📐 設計風壓計算式</div>
    <div class="rc-loc"><strong>${data.buildingType}</strong><span class="slash">／</span>${data.enclosureType}</div>
    <div class="rc-loc" style="margin-top:3px;font-size:10px;color:var(--muted);font-weight:400">${data.designSystem}</div>
    <div class="rc-loc" style="margin-top:2px;font-size:10px;color:var(--muted);font-weight:400">${data.subcategory}</div>
    ${mainRows ? `
    <div style="font-size:8px;letter-spacing:.1em;color:var(--muted);margin:8px 0 3px;font-family:'DM Mono',monospace">FORMULAS</div>
    <table style="width:100%;border-collapse:collapse">${mainRows}</table>` : ''}
    ${extraRows ? `
    <div style="font-size:8px;letter-spacing:.1em;color:var(--muted);margin:10px 0 3px;font-family:'DM Mono',monospace;border-top:1px dashed var(--border);padding-top:8px">相關公式（風速壓）</div>
    <table style="width:100%;border-collapse:collapse">${extraRows}</table>` : ''}
    <div class="rc-time">${formatTime(data.timestamp)}</div>`;
  replaceCard('f', card);
  // Relay formula selection to S-section iframe
  const sf2 = document.getElementById('frame-s');
  if (sf2 && sf2.contentWindow) sf2.contentWindow.postMessage(data, '*');
}


// ══════════════════════════════════════════════════════════════════
//  S 區通訊：傳遞暫存結果給 frame-s
// ══════════════════════════════════════════════════════════════════

function getResultByType(type) {
  return HubState.get(type);
}

/** Push all cached results as a type→data map to the S-section iframe */
function pushResultsToSecS() {
  const sf = document.getElementById('frame-s');
  if (!sf || !sf.contentWindow) return;
  sf.contentWindow.postMessage({ source: 'hub_to_sec_s_results', results: HubState.getAll() }, '*');
}

function setPill(k, text) {
  const el = document.getElementById('pill-'+k);
  if (!el) return;
  el.textContent = text; el.style.display = 'inline-block';
}

function replaceCard(section, card) {
  // Remove previous card from the same section (DOM only; HubState already updated by caller)
  const prev = document.querySelector('.result-card[data-section="'+section+'"]');
  if (prev) prev.remove();

  // Insert in A→F order: find the first existing card whose section > current section
  const order = ['a','b','c','d','e','f','g','s','r'];
  const body = document.getElementById('results-body');
  const currentRank = order.indexOf(section);
  let insertBefore = null;
  const existing = body.querySelectorAll('.result-card[data-section]');
  for (const el of existing) {
    const elRank = order.indexOf(el.dataset.section);
    if (elRank > currentRank) { insertBefore = el; break; }
  }
  if (insertBefore) {
    body.insertBefore(card, insertBefore);
  } else {
    body.appendChild(card);
  }

  card.scrollIntoView({behavior:'smooth', block:'nearest'});
  const cnt = document.getElementById('result-count');
  cnt.style.transform='scale(1.5)';
  setTimeout(()=>cnt.style.transform='scale(1)',200);
}

function appendCard(card) {
  // Legacy — kept for safety; most callers now use replaceCard
  const body = document.getElementById('results-body');
  body.appendChild(card);
  card.scrollIntoView({behavior:'smooth', block:'nearest'});
  const cnt = document.getElementById('result-count');
  cnt.style.transform='scale(1.5)';
  setTimeout(()=>cnt.style.transform='scale(1)',200);
}

function removeCard(id, type) {
  const el = document.getElementById(id); if (!el) return;
  // Resolve type from argument or data attribute
  const resolvedType = type || (el && el.dataset.type) || null;
  el.style.opacity='0'; el.style.transform='scale(0.9)'; el.style.transition='all 0.2s';
  setTimeout(()=>{
    el.remove();
    if (resolvedType) HubState.remove(resolvedType);
    if (Object.keys(HubState.getAll()).length === 0) showEmpty();
  },200);
}

function clearAll() {
  HubState.clear();
  document.getElementById('results-body').innerHTML='<div class="empty-hint" id="empty-hint"><span class="empty-dot"></span>請展開 A、B 或 C 區進行查詢，結果將自動顯示於此</div>';
  updateCount();
}

function updateCount() {
  document.getElementById('result-count').textContent = Object.keys(HubState.getAll()).length;
}
function hideEmpty(){ const h=document.getElementById('empty-hint'); if(h)h.style.display='none'; }
function showEmpty(){ document.getElementById('results-body').innerHTML='<div class="empty-hint" id="empty-hint"><span class="empty-dot"></span>請展開 A、B 或 C 區進行查詢，結果將自動顯示於此</div>'; }
function formatTime(iso){
  try{ const d=new Date(iso),p=n=>String(n).padStart(2,'0');
    return `${d.getFullYear()}-${p(d.getMonth()+1)}-${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}:${p(d.getSeconds())}`;
  }catch{return '';}
}
