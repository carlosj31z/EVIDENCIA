/* FASC-205-03 · Evidencia visual de acciones de Controles de Cambio
   - Datos: Supabase (tablas cc_acciones / cc_evidencias, bucket "evidencias") o modo local en memoria.
   - Genera el .xlsx a partir de la plantilla original (assets/template.js) conservando formato, encabezado y pie. */
(() => {
'use strict';

/* ---------- Utilidades ---------- */
const $ = s => document.querySelector(s);
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const xesc = s => String(s ?? '').replace(/[&<>"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c])).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g,'');
const clean = v => v == null ? '' : String(v).replace(/ /g,' ').replace(/[ \t]+\n/g,'\n').replace(/ {2,}/g,' ').trim();
const norm = s => String(s ?? '').normalize('NFD').replace(/[̀-ͯ]/g,'').toUpperCase().replace(/\s+/g,' ').trim();
const uid = () => (crypto.randomUUID ? crypto.randomUUID() : Date.now().toString(36)+Math.random().toString(36).slice(2));
const fmtDate = iso => { if(!iso) return ''; const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(iso); return m ? `${m[3]}/${m[2]}/${m[1]}` : iso; };
const collator = new Intl.Collator('es', {numeric:true, sensitivity:'base'});
function toast(msg, ms=2600){ const t=$('#toast'); t.textContent=msg; t.hidden=false; clearTimeout(toast._t); toast._t=setTimeout(()=>t.hidden=true, ms); }
function notice(msg, kind=''){ const n=$('#notice'); if(!msg){n.hidden=true;return;} n.className='notice '+kind; n.textContent=msg; n.hidden=false; }
const debounce = (fn, ms) => { let t; return (...a) => { clearTimeout(t); t = setTimeout(() => fn(...a), ms); }; };

/* ---------- Capa de datos ---------- */
const CFG = window.APP_CONFIG || {};
const BUCKET = CFG.BUCKET || 'evidencias';
const sb = (CFG.SUPABASE_URL && CFG.SUPABASE_ANON_KEY && window.supabase) ? window.supabase.createClient(CFG.SUPABASE_URL, CFG.SUPABASE_ANON_KEY) : null;
const blobCache = new Map(); // path -> Blob
const DATA_COLS = ['codigo_cc','num_accion','descripcion','responsable','area','fecha','status_acciones','status','comentario','fecha_evidencia_ssuit'];

const Cloud = {
  async loadAll(){
    const a = await sb.from('cc_acciones').select('*').range(0, 9999);
    if (a.error) throw a.error;
    const e = await sb.from('cc_evidencias').select('*').order('orden').range(0, 99999);
    if (e.error) throw e.error;
    return {acciones:a.data, evidencias:e.data};
  },
  async upsertAcciones(rows){
    const out=[];
    for (let i=0;i<rows.length;i+=400){
      const r = await sb.from('cc_acciones').upsert(rows.slice(i,i+400), {onConflict:'codigo_cc,num_accion'}).select('*');
      if (r.error) throw r.error; out.push(...r.data);
    }
    return out;
  },
  async setInactive(ids){ if(!ids.length) return; const r=await sb.from('cc_acciones').update({activo:false}).in('id',ids); if(r.error) throw r.error; },
  async updateAccion(id, patch){ const r=await sb.from('cc_acciones').update({...patch, updated_at:new Date().toISOString()}).eq('id',id); if(r.error) throw r.error; },
  async upload(path, blob){ const r=await sb.storage.from(BUCKET).upload(path, blob, {contentType:blob.type, upsert:false}); if(r.error) throw r.error; blobCache.set(path, blob); },
  async removeFiles(paths){ paths=paths.filter(Boolean); if(!paths.length) return; await sb.storage.from(BUCKET).remove(paths); paths.forEach(p=>blobCache.delete(p)); },
  async insertEvid(row){ const r=await sb.from('cc_evidencias').insert(row).select('*').single(); if(r.error) throw r.error; return r.data; },
  async updateEvid(id, patch){ const r=await sb.from('cc_evidencias').update(patch).eq('id',id); if(r.error) throw r.error; },
  async deleteEvid(id){ const r=await sb.from('cc_evidencias').delete().eq('id',id); if(r.error) throw r.error; },
  async blob(path){ if(blobCache.has(path)) return blobCache.get(path); const r=await sb.storage.from(BUCKET).download(path); if(r.error) throw r.error; blobCache.set(path, r.data); return r.data; }
};
const Local = {
  rows:[], ev:[],
  async loadAll(){ return {acciones:this.rows, evidencias:this.ev}; },
  async upsertAcciones(rows){
    return rows.map(r => { let ex=this.rows.find(x=>x.codigo_cc===r.codigo_cc && x.num_accion===r.num_accion);
      if (ex) Object.assign(ex, r); else { ex={id:uid(), planta:2, anexo:'-', seccion:null, texto_accion:null, por_pagina:null, generado_at:null, ...r}; this.rows.push(ex); }
      return {...ex}; });
  },
  async setInactive(ids){ this.rows.forEach(r=>{ if(ids.includes(r.id)) r.activo=false; }); },
  async updateAccion(id, patch){ Object.assign(this.rows.find(r=>r.id===id)||{}, patch); },
  async upload(path, blob){ blobCache.set(path, blob); },
  async removeFiles(paths){ paths.forEach(p=>blobCache.delete(p)); },
  async insertEvid(row){ const r={id:uid(), ...row}; this.ev.push(r); return {...r}; },
  async updateEvid(id, patch){ Object.assign(this.ev.find(e=>e.id===id)||{}, patch); },
  async deleteEvid(id){ this.ev=this.ev.filter(e=>e.id!==id); },
  async blob(path){ return blobCache.get(path); }
};
const DB = sb ? Cloud : Local;

/* ---------- Estado ---------- */
let acciones = [];            // filas de cc_acciones
let evid = new Map();         // accion_id -> [evidencias ordenadas]
let filters = {};             // colKey -> Set(claves normalizadas permitidas)
let sort = {k:null, dir:1};
let query = '';
let current = null;           // acción abierta en el panel
const thumbURL = new Map();   // path -> objectURL

/* ---------- Columnas ---------- */
const BLANK = '(Vacías)';
const evState = a => a.generado_at ? 'Descargado' : ((evid.get(a.id)||[]).length ? 'Con imágenes' : 'Sin imágenes');
const COLS = [
  {k:'codigo_cc', t:'Código CC', cls:'mono'},
  {k:'num_accion', t:'N° acción', num:true, cls:'mono'},
  {k:'descripcion', t:'Descripción de la acción', cls:'c-desc'},
  {k:'responsable', t:'Responsable'},
  {k:'area', t:'Área'},
  {k:'fecha', t:'Fecha', get:a=>fmtDate(a.fecha), sortv:a=>a.fecha||'', cls:'mono'},
  {k:'status_acciones', t:'Status de acciones', pill:true},
  {k:'status', t:'Status', pill:true},
  {k:'comentario', t:'Comentario', cls:'c-com'},
  {k:'_ev', t:'Evidencia', get:evState, pill:true},
];
const cellVal = (a,c) => { const v = c.get ? c.get(a) : a[c.k]; const s = clean(v).replace(/\n/g,' '); return s || BLANK; };
const keyOf = s => String(s).toLowerCase().trim();
function pillClass(v){
  const s = norm(v);
  if (/REALIZADO|DESCARGADO|CERRADO/.test(s)) return 'p-ok';
  if (/VENCIDO|ANULAR/.test(s)) return 'p-bad';
  if (/PENDIENTE|SIN IMAGENES/.test(s)) return 'p-warn';
  if (/PROCESO|PLAZO|CON IMAGENES/.test(s)) return 'p-info';
  return 'p-neu';
}

/* ---------- Filtros / orden ---------- */
function passes(a, exceptKey){
  if (!$('#showInactive').checked && a.activo === false) return false;
  for (const c of COLS){ if (c.k===exceptKey) continue; const f=filters[c.k]; if (f && !f.has(keyOf(cellVal(a,c)))) return false; }
  if (query){ const q=norm(query); if (!COLS.some(c => norm(cellVal(a,c)).includes(q))) return false; }
  return true;
}
function visibleRows(){
  const rows = acciones.filter(a=>passes(a));
  const cmpDefault = (x,y) => collator.compare(x.codigo_cc,y.codigo_cc) || (parseFloat(x.num_accion)||0)-(parseFloat(y.num_accion)||0);
  if (!sort.k) return rows.sort(cmpDefault);
  const c = COLS.find(c=>c.k===sort.k);
  const sv = a => c.sortv ? c.sortv(a) : c.num ? (parseFloat(a[c.k])||0) : cellVal(a,c);
  return rows.sort((x,y)=>{ const A=sv(x),B=sv(y); const r = typeof A==='number' ? A-B : collator.compare(A,B); return (r||cmpDefault(x,y))*(r?sort.dir:1); });
}

/* ---------- Render tabla ---------- */
const FUNNEL = '<svg viewBox="0 0 24 24" stroke-width="2"><path d="M3 5h18l-7 8v6l-4-2v-4z"/></svg>';
const CARET = '<svg viewBox="0 0 24 24" stroke-width="2"><path d="m6 9 6 6 6-6"/></svg>';
function renderHead(){
  $('#thead').innerHTML = COLS.map(c => {
    const on = !!filters[c.k] || sort.k===c.k;
    const arrow = sort.k===c.k ? (sort.dir>0?' ↑':' ↓') : '';
    return `<th><div class="th-in"><span>${esc(c.t)}${arrow}</span><button type="button" class="fbtn${filters[c.k]?' on':''}" data-col="${c.k}" aria-label="Filtrar ${esc(c.t)}">${filters[c.k]?FUNNEL:CARET}</button></div></th>`;
  }).join('');
}
function renderTable(){
  renderHead();
  const rows = visibleRows();
  const tb = $('#tbody');
  tb.innerHTML = rows.map(a => '<tr data-id="'+a.id+'" class="'+(a.activo===false?'inactive ':'')+(current&&current.id===a.id?'sel':'')+'">' + COLS.map(c => {
    let v = c.get ? c.get(a) : a[c.k]; v = clean(v);
    if (c.k==='_ev'){ const n=(evid.get(a.id)||[]).length; return `<td><span class="pill ${pillClass(v)}">${esc(v)}</span>${n?` <span class="hint mono">${n} img</span>`:''}</td>`; }
    if (c.pill && v) return `<td><span class="pill ${pillClass(v)}">${esc(v)}</span></td>`;
    if (c.k==='descripcion') return `<td class="${c.cls}"><div>${esc(v)}</div></td>`;
    return `<td class="${c.cls||''}">${esc(v).replace(/\n/g,'<br>')}</td>`;
  }).join('') + '</tr>').join('');
  const total = acciones.filter(a=>$('#showInactive').checked || a.activo!==false).length;
  $('#count').innerHTML = `<strong>${rows.length}</strong> de ${total} acciones`;
  const em = $('#empty');
  if (!acciones.length){ em.innerHTML = '<b>Aún no hay acciones cargadas</b>Usa “Importar Excel de CC” y elige el archivo de controles de cambio pendientes.'; em.hidden=false; }
  else if (!rows.length){ em.innerHTML = '<b>Ninguna acción coincide con los filtros</b>Borra o ajusta los filtros para ver más filas.'; em.hidden=false; }
  else em.hidden = true;
  // chips
  const chips = COLS.filter(c=>filters[c.k]).map(c=>`<span class="chip">${esc(c.t)}: ${filters[c.k].size} valor${filters[c.k].size===1?'':'es'}<button type="button" data-clear="${c.k}" aria-label="Quitar filtro">×</button></span>`);
  $('#chips').innerHTML = chips.join('');
  $('#btnClear').hidden = !chips.length && !query && !sort.k;
}

/* ---------- Popover de filtro tipo Excel ---------- */
let pop = null;
function closePop(){ if(pop){ pop.remove(); pop=null; } }
function openFilter(colKey, btn){
  closePop();
  const c = COLS.find(x=>x.k===colKey);
  const base = acciones.filter(a=>passes(a, colKey));
  const map = new Map();
  for (const a of base){ const v=cellVal(a,c), k=keyOf(v); if(!map.has(k)) map.set(k,{k,label:v,n:0}); map.get(k).n++; }
  const items = [...map.values()].sort((x,y)=> x.label===BLANK?1 : y.label===BLANK?-1 : (c.num ? (parseFloat(x.label)||0)-(parseFloat(y.label)||0) : c.sortv ? collator.compare(x.label.split('/').reverse().join(''), y.label.split('/').reverse().join('')) : collator.compare(x.label,y.label)));
  const cur = filters[colKey];
  const checked = new Set(items.filter(i=>!cur || cur.has(i.k)).map(i=>i.k));
  pop = document.createElement('div'); pop.className='pop'; pop.setAttribute('role','dialog');
  pop.innerHTML = `<div class="sorts">
      <button type="button" data-s="1" class="${sort.k===colKey&&sort.dir>0?'cur':''}">Ordenar de A a Z</button>
      <button type="button" data-s="-1" class="${sort.k===colKey&&sort.dir<0?'cur':''}">Ordenar de Z a A</button>
      <button type="button" data-s="clear" ${filters[colKey]?'':'disabled'}>Borrar filtro de “${esc(c.t)}”</button></div>
    <div class="fsearch"><input type="search" id="popq" placeholder="Buscar" autocomplete="off"></div>
    <div class="flist" id="flist"></div>
    <div class="fact"><button type="button" class="btn" data-a="cancel">Cancelar</button><button type="button" class="btn primary" data-a="ok">Aceptar</button></div>`;
  document.body.appendChild(pop);
  const list = pop.querySelector('#flist'), qi = pop.querySelector('#popq');
  const shown = () => { const q=norm(qi.value); return items.filter(i=>!q || norm(i.label).includes(q)); };
  function draw(){
    const vis = shown(); const allOn = vis.length && vis.every(i=>checked.has(i.k));
    list.innerHTML = `<label><input type="checkbox" data-all ${allOn?'checked':''}><span><b>(Seleccionar todo${qi.value?' lo buscado':''})</b></span></label>` +
      vis.map((i,ix)=>`<label title="${esc(i.label)}"><input type="checkbox" data-i="${ix}" ${checked.has(i.k)?'checked':''}><span>${esc(i.label)}</span><em>${i.n}</em></label>`).join('');
    list.onchange = e => { const t=e.target; if (t.dataset.all!==undefined){ vis.forEach(i=> t.checked?checked.add(i.k):checked.delete(i.k)); } else { const it=vis[+t.dataset.i]; t.checked?checked.add(it.k):checked.delete(it.k); } draw(); };
  }
  draw(); qi.oninput = draw;
  pop.onclick = e => {
    const s = e.target.closest('[data-s]'), a = e.target.closest('[data-a]');
    if (s){ if (s.dataset.s==='clear'){ delete filters[colKey]; } else { sort = {k:colKey, dir:+s.dataset.s}; } closePop(); renderTable(); }
    if (a){ if (a.dataset.a==='ok'){
        let sel = qi.value ? new Set(shown().filter(i=>checked.has(i.k)).map(i=>i.k)) : new Set(checked);
        if (!qi.value && sel.size===items.length) delete filters[colKey]; else filters[colKey]=sel;
        renderTable(); }
      closePop(); }
  };
  // posición
  const r = btn.getBoundingClientRect(), w = pop.offsetWidth, h = pop.offsetHeight;
  let left = Math.min(r.left, innerWidth - w - 12); left = Math.max(12, left);
  let top = r.bottom + 4; if (top + h > innerHeight - 12) top = Math.max(12, innerHeight - h - 12);
  pop.style.left = left+'px'; pop.style.top = top+'px';
  qi.focus();
}
document.addEventListener('mousedown', e => { if (pop && !pop.contains(e.target) && !e.target.closest('.fbtn')) closePop(); });
document.addEventListener('keydown', e => { if (e.key==='Escape'){ if (pop) closePop(); else if (!$('#editor').hidden) Editor.close(); else if (!$('#drawer').hidden) closeDrawer(); } });

/* ---------- Importación del Excel de controles de cambio ---------- */
function parseCC(ab){
  const wb = XLSX.read(ab, {type:'array'});
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rg = XLSX.utils.decode_range(ws['!ref']);
  const grid = [];
  for (let r=rg.s.r; r<=rg.e.r; r++){ const row=[]; for (let c=rg.s.c; c<=rg.e.c; c++) row.push(ws[XLSX.utils.encode_cell({r,c})] || null); grid.push(row); }
  for (const m of ws['!merges']||[]){ const tl = grid[m.s.r-rg.s.r]?.[m.s.c-rg.s.c]; for (let r=m.s.r;r<=m.e.r;r++) for (let c=m.s.c;c<=m.e.c;c++) grid[r-rg.s.r][c-rg.s.c] = tl; }
  const hi = grid.findIndex(row => row.some(c => c && norm(c.v)==='CODIGO CC'));
  if (hi<0) throw new Error('No se encontró la fila de encabezados (columna “CÓDIGO CC”).');
  const H = grid[hi].map(c => c ? norm(c.v) : '');
  const col = {};
  H.forEach((h,i)=>{
    if (h==='CODIGO CC') col.codigo_cc=i;
    else if (/^N.? ?DE ACCION|^N° ?ACCION|^NO? ?ACCION/.test(h) || (h.startsWith('N') && h.includes('ACCION') && !h.includes('DESCRIP') && !h.includes('STATUS'))) col.num_accion ??= i;
    else if (h.startsWith('DESCRIPCION')) col.descripcion=i;
    else if (h.startsWith('RESPONSABLE')) col.responsable=i;
    else if (h.startsWith('AREA')) col.area=i;
    else if (h.startsWith('FECHA EVIDENCIA')) col.fecha_evidencia_ssuit=i;
    else if (h==='FECHA') col.fecha=i;
    else if (h.startsWith('STATUS DE ACCION')) col.status_acciones=i;
    else if (h==='STATUS') col.status=i;
    else if (h.startsWith('COMENTARIO')) col.comentario=i;
  });
  if (col.codigo_cc==null || col.num_accion==null || col.descripcion==null) throw new Error('Faltan columnas obligatorias: CÓDIGO CC, Nº DE ACCIÓN o DESCRIPCIÓN DE LA ACCIÓN.');
  const dateOf = cell => {
    if (!cell) return '';
    if (cell.t==='n'){ const d = XLSX.SSF.parse_date_code(cell.v); if (d) return `${d.y}-${String(d.m).padStart(2,'0')}-${String(d.d).padStart(2,'0')}`; }
    const m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})/.exec(clean(cell.w||cell.v)); if (m) return `${m[3]}-${m[2].padStart(2,'0')}-${m[1].padStart(2,'0')}`;
    return '';
  };
  const txt = (row,k) => col[k]==null ? '' : clean(row[col[k]] ? (row[col[k]].w ?? row[col[k]].v) : '');
  const addUnique = (base, add, sep) => { if (!add) return base; if (!base) return add; return base.split(sep).map(s=>s.trim()).includes(add.trim()) || base.includes(add) ? base : base + sep + add; };
  const out = []; let last = null;
  for (let r=hi+1; r<grid.length; r++){
    const row = grid[r];
    let cod = txt(row,'codigo_cc'), num = txt(row,'num_accion');
    const desc = col.descripcion!=null && row[col.descripcion] ? clean(row[col.descripcion].v) : '';
    const resp = txt(row,'responsable'), com = txt(row,'comentario');
    if (!cod && !num && !desc && !resp) continue;
    if (!cod && num && last) cod = last.codigo_cc;
    const isCont = last && ((!cod && !num) || (cod===last.codigo_cc && num===last.num_accion));
    if (isCont){
      last.descripcion = addUnique(last.descripcion, desc, '\n');
      resp.split('\n').forEach(p => last.responsable = addUnique(last.responsable, p.trim(), '\n'));
      last.comentario = addUnique(last.comentario, com, ' · ');
      if (!last.status) last.status = txt(row,'status');
      continue;
    }
    if (!cod) continue;
    last = { codigo_cc:cod, num_accion:num, descripcion:desc, responsable:resp, area:txt(row,'area'),
      fecha: col.fecha!=null ? dateOf(row[col.fecha]) : '', status_acciones:txt(row,'status_acciones'), status:txt(row,'status'),
      comentario:com, fecha_evidencia_ssuit: col.fecha_evidencia_ssuit!=null ? (dateOf(row[col.fecha_evidencia_ssuit]) || txt(row,'fecha_evidencia_ssuit')) : '' };
    out.push(last);
  }
  return out;
}
async function importFile(file){
  try{
    notice('Leyendo '+file.name+'…');
    const rows = parseCC(await file.arrayBuffer());
    if (!rows.length) throw new Error('El archivo no tiene acciones.');
    // Mismo código + acción repetido en el Excel: se conserva la primera fila
    const seen = new Set(), uniq = rows.filter(r=>{ const k=r.codigo_cc+'|'+r.num_accion; if(seen.has(k)) return false; seen.add(k); return true; });
    const payload = uniq.map(r => { const o={activo:true}; DATA_COLS.forEach(k=>o[k]= r[k]==='' ? null : r[k]); return o; });
    const before = new Set(acciones.map(a=>a.codigo_cc+'|'+a.num_accion));
    const saved = await DB.upsertAcciones(payload);
    const keys = new Set(saved.map(s=>s.codigo_cc+'|'+s.num_accion));
    const gone = acciones.filter(a => a.activo!==false && !keys.has(a.codigo_cc+'|'+a.num_accion)).map(a=>a.id);
    await DB.setInactive(gone);
    await reload();
    const nuevos = saved.filter(s=>!before.has(s.codigo_cc+'|'+s.num_accion)).length;
    notice(`Importadas ${saved.length} acciones de “${file.name}”: ${nuevos} nuevas, ${saved.length-nuevos} actualizadas${gone.length?`, ${gone.length} ya no figuran (ocultas)`:''}.`, 'ok');
  } catch(err){ console.error(err); notice('No se pudo importar: '+(err.message||err), 'err'); }
}

/* ---------- Carga ---------- */
async function reload(){
  const {acciones:a, evidencias:e} = await DB.loadAll();
  acciones = a.map(x=>({...x}));
  evid = new Map();
  for (const x of e){ if(!evid.has(x.accion_id)) evid.set(x.accion_id, []); evid.get(x.accion_id).push({...x}); }
  evid.forEach(l=>l.sort((p,q)=>p.orden-q.orden));
  if (current) current = acciones.find(x=>x.id===current.id) || null;
  renderTable();
}

/* ============================================================
   Geometría del formato (en píxeles de Excel, Calibri 11 → 7 px por carácter)
   ============================================================ */
const COLPX = [75,75,22,115,75,75,25,16,19,24,76,17];   // anchos reales de A..L en la plantilla
const SHEET_W = COLPX.reduce((s,v)=>s+v,0);              // 614 px
const HEAD_PT = [15.75,17.25,15.75,6,9.75,15,6.75,15,8.25,7.5,15,15,15,8.25,6.75,39.75,12]; // filas 1..17
const BLOCK_ROWS = 34;                                     // filas 18..51 = bloque de evidencias (1 página)
const pt2px = pt => Math.round(pt*4/3);
const q75 = pt => Math.round(pt/0.75)*0.75;

function accionTexto(a){ return a.texto_accion ?? `Acción N° ${a.num_accion}:  ${a.descripcion||''}`; }
function seccionDe(a){ return a.seccion ?? (a.area || ''); }
function fileName(a){ return `FASC-205-03 EVIDENCIA VISUAL ${a.codigo_cc} ACCION ${a.num_accion}.xlsx`.replace(/[\\/:*?"<>|\r\n]+/g,' ').replace(/\s{2,}/g,' '); }

/* Altura de la fila 16 según el largo del texto; lo que crece se descuenta del bloque de evidencias para que siga entrando en una hoja A4. */
function heights(texto){
  const CPL = 84; let lines = 0;
  for (const p of String(texto).split('\n')) lines += Math.max(1, Math.ceil(p.length / CPL));
  const total = Math.max(51.75, q75(lines*15 + 9));
  const r16 = total - 12;
  const block = [15, ...Array(31).fill(15), 52.5, 0.75];   // filas 18, 19..49, 50, 51
  let excess = r16 - 39.75, i = 32;
  while (excess > 0 && i >= 1){ const min = i===32 ? 6 : 0.75; const d = Math.min(block[i]-min, excess); block[i] = q75(block[i]-d); excess -= d; i--; }
  return {r16, block};
}
/* Distribuye las imágenes en páginas y las centra en su espacio. Coordenadas relativas al borde izq. de la col A y al borde sup. de la fila 18 del bloque. */
function computeLayout(a, imgs){
  const {r16, block} = heights(accionTexto(a));
  const rowPx = block.map(pt2px);
  const PAD = 8, GAP = 10;
  const box = {x:PAD, y:PAD, w:SHEET_W - 2*PAD, h:rowPx.slice(0,33).reduce((s,v)=>s+v,0) - 2*PAD};
  const n = imgs.length;
  const per = a.por_pagina ? +a.por_pagina : (n<=1 ? 1 : n===2 ? 2 : 4);
  const slots = per===1 ? [[0,0,1,1]] : per===2 ? [[0,0,1,.5],[0,.5,1,.5]] : [[0,0,.5,.5],[.5,0,.5,.5],[0,.5,.5,.5],[.5,.5,.5,.5]];
  const pages = [];
  for (let i=0;i<Math.max(n,1);i+=per){
    const pg = [];
    imgs.slice(i,i+per).forEach((im,j)=>{
      const [sx,sy,sw,sh] = slots[j];
      const slot = {x: box.x + sx*box.w + (sx>0?GAP/2:0), y: box.y + sy*box.h + (sy>0?GAP/2:0), w: sw*box.w - (sw<1?GAP/2:0), h: sh*box.h - (sh<1?GAP/2:0)};
      const s = Math.min(slot.w/im.ancho, slot.h/im.alto);
      const w = Math.floor(im.ancho*s), h = Math.floor(im.alto*s);
      pg.push({im, x: Math.round(slot.x + (slot.w-w)/2), y: Math.round(slot.y + (slot.h-h)/2), w, h});
    });
    pages.push(pg);
  }
  return {r16, block, rowPx, box, pages, per};
}

/* ============================================================
   Generación del XLSX a partir de la plantilla
   ============================================================ */
const siXml = t => `<si><t xml:space="preserve">${xesc(t)}</t></si>`;
function pxToCol(x){ let acc=0; for (let c=0;c<COLPX.length;c++){ if (x < acc+COLPX[c] || c===COLPX.length-1) return [c, Math.max(0, x-acc)]; acc+=COLPX[c]; } }
function pxToRow(y, rowPx, base){ let acc=0; for (let i=0;i<rowPx.length;i++){ if (y < acc+rowPx[i] || i===rowPx.length-1) return [base+i, Math.max(0,y-acc)]; acc+=rowPx[i]; } }
const EMU = 9525;

async function buildXlsx(a, imgs){
  const L = computeLayout(a, imgs);
  const nPages = L.pages.length;
  const zip = await JSZip.loadAsync(window.FASC_TEMPLATE_B64, {base64:true});

  // Textos (sharedStrings: 11 = anexo, 13 = código, 14 = sección, 15 = acción)
  let ss = await zip.file('xl/sharedStrings.xml').async('string'); let k = 0;
  ss = ss.replace(/<si>[\s\S]*?<\/si>/g, m => { const i=k++; return i===11 ? siXml(a.anexo||'-') : i===13 ? siXml(a.codigo_cc) : i===14 ? siXml(seccionDe(a)) : i===15 ? siXml(accionTexto(a)) : m; });
  zip.file('xl/sharedStrings.xml', ss);

  // Hoja
  let sh = await zip.file('xl/worksheets/sheet1.xml').async('string');
  const planta = +a.planta || 2;
  sh = sh.replace(/<c r="J6" s="3" t="s"><v>\d+<\/v><\/c>/, `<c r="J6" s="3" t="s"><v>${planta===1?0:10}</v></c>`)
         .replace(/<c r="J8" s="3" t="s"><v>\d+<\/v><\/c>/, `<c r="J8" s="3" t="s"><v>${planta===2?0:10}</v></c>`);
  const rows = {}; sh.replace(/<row r="(\d+)"[\s\S]*?<\/row>/g, (m, r) => { rows[+r] = m; return m; });
  const setHt = (xml, pt) => { pt = +pt.toFixed(2);
    if (/<row [^>]*\bht="/.test(xml)) return xml.replace(/(<row [^>]*?)\bht="[\d.]+"( customHeight="1")?/, `$1ht="${pt}" customHeight="1"`);
    return pt===15 ? xml : xml.replace(/<row r="(\d+)"/, `<row r="$1" ht="${pt}" customHeight="1"`); };
  const renum = (xml, nr) => xml.replace(/<row r="\d+"/, `<row r="${nr}"`).replace(/ r="([A-Z]+)\d+"/g, ` r="$1${nr}"`);
  rows[16] = setHt(rows[16], L.r16);
  let data = '';
  for (let r=1;r<=17;r++) data += rows[r];
  for (let p=0;p<nPages;p++) for (let i=0;i<BLOCK_ROWS;i++) data += renum(setHt(rows[18+i], L.block[i]), 18+i+p*BLOCK_ROWS);
  const lastRow = 17 + nPages*BLOCK_ROWS;
  sh = sh.replace(/<sheetData>[\s\S]*<\/sheetData>/, `<sheetData>${data}</sheetData>`)
         .replace(/<dimension ref="[^"]+"\/>/, `<dimension ref="A1:L${lastRow}"/>`);
  let merges = (sh.match(/<mergeCell ref="[^"]+"\/>/g) || []);
  for (let p=1;p<nPages;p++){ const r=18+p*BLOCK_ROWS; merges.push(`<mergeCell ref="A${r}:L${r}"/>`); }
  sh = sh.replace(/<mergeCells count="\d+">[\s\S]*?<\/mergeCells>/, `<mergeCells count="${merges.length}">${merges.join('')}</mergeCells>`);
  let extra = '';
  if (nPages>1){ const brks=[]; for (let p=1;p<nPages;p++) brks.push(`<brk id="${17+p*BLOCK_ROWS}" max="16383" man="1"/>`); extra += `<rowBreaks count="${brks.length}" manualBreakCount="${brks.length}">${brks.join('')}</rowBreaks>`; }
  if (imgs.length) extra += '<drawing r:id="rIdEvid"/>';
  sh = sh.replace('<legacyDrawingHF', extra + '<legacyDrawingHF');
  zip.file('xl/worksheets/sheet1.xml', sh);

  // Área de impresión (+ títulos a imprimir 1:17 cuando hay varias páginas)
  let wbx = await zip.file('xl/workbook.xml').async('string');
  wbx = wbx.replace(/\$A\$1:\$L\$51/, () => '$A$1:$L$' + lastRow);
  if (nPages>1) wbx = wbx.replace(/(<definedName name="_xlnm.Print_Area"[^>]*>[^<]*<\/definedName>)/, m => m + `<definedName name="_xlnm.Print_Titles" localSheetId="0">'EVIDENCIA '!$1:$17</definedName>`);
  zip.file('xl/workbook.xml', wbx);

  // Imágenes
  if (imgs.length){
    let anchors = '', rels = '', id = 1;
    const exts = new Set();
    for (let p=0;p<nPages;p++){
      for (const it of L.pages[p]){
        const blob = await DB.blob(it.im.storage_path);
        const ext = /png/.test(blob.type) ? 'png' : 'jpeg'; exts.add(ext);
        zip.file(`xl/media/evidencia${id}.${ext}`, await blob.arrayBuffer());
        rels += `<Relationship Id="rId${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/evidencia${id}.${ext}"/>`;
        const [c, co] = pxToCol(it.x);
        const [r, ro] = pxToRow(it.y, L.rowPx, 17 + p*BLOCK_ROWS);
        anchors += `<xdr:oneCellAnchor><xdr:from><xdr:col>${c}</xdr:col><xdr:colOff>${Math.round(co*EMU)}</xdr:colOff><xdr:row>${r}</xdr:row><xdr:rowOff>${Math.round(ro*EMU)}</xdr:rowOff></xdr:from><xdr:ext cx="${it.w*EMU}" cy="${it.h*EMU}"/>`+
          `<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${id+1}" name="Evidencia ${id}"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>`+
          `<xdr:blipFill><a:blip r:embed="rId${id}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>`+
          `<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${it.w*EMU}" cy="${it.h*EMU}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>`;
        id++;
      }
    }
    zip.file('xl/drawings/drawing1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${anchors}</xdr:wsDr>`);
    zip.file('xl/drawings/_rels/drawing1.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships>`);
    let srels = await zip.file('xl/worksheets/_rels/sheet1.xml.rels').async('string');
    srels = srels.replace('</Relationships>', `<Relationship Id="rIdEvid" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>`);
    zip.file('xl/worksheets/_rels/sheet1.xml.rels', srels);
    let ct = await zip.file('[Content_Types].xml').async('string');
    if (exts.has('png') && !/Extension="png"/.test(ct)) ct = ct.replace('<Default Extension="rels"', '<Default Extension="png" ContentType="image/png"/><Default Extension="rels"');
    ct = ct.replace('</Types>', '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>');
    zip.file('[Content_Types].xml', ct);
  }
  return zip.generateAsync({type:'blob', mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', compression:'DEFLATE'});
}

/* ============================================================
   Vista previa (réplica de la hoja A4 impresa)
   ============================================================ */
const COLX = COLPX.reduce((acc,w,i)=>{ acc.push(acc[i]+w); return acc; }, [0]);
async function renderPreview(){
  if (!current) return;
  const a = current, imgs = evid.get(a.id) || [];
  const L = computeLayout(a, imgs);
  const headPx = HEAD_PT.map(pt2px); headPx[15] = pt2px(L.r16);
  const ROWY = headPx.reduce((acc,h,i)=>{ acc.push(acc[i]+h); return acc; }, [0]);  // y de inicio de cada fila 1..18
  const blockH = L.rowPx.reduce((s,v)=>s+v,0);
  const SCALE = 0.95, PAGE_W = 794, LEFT = Math.round((PAGE_W - SHEET_W*SCALE)/2), TOP = 102;
  const cell = (c1, r1, c2, r2, html, cls='', style='') => {
    const x = COLX[c1], y = ROWY[r1-1], w = COLX[c2+1]-x, h = ROWY[r2]-y;
    return `<div class="c ${cls}" style="left:${x}px;top:${y}px;width:${w}px;height:${h}px;${style}">${html}</div>`;
  };
  const head = [
    `<div class="frame" style="left:0;top:${ROWY[1]}px;width:${SHEET_W}px;height:${ROWY[3]-ROWY[1]}px"></div>`,
    cell(0,2,4,3,'<b style="font-size:16px;width:100%;text-align:center">EVIDENCIA VISUAL</b>'),
    cell(5,2,11,2,'<b style="width:100%;text-align:center;font-size:13px">CÓDIGO</b>','grey','border:1px solid #000'),
    cell(5,3,11,3,`<span style="width:100%;text-align:center;font-size:14.5px">${esc(a.codigo_cc)}</span>`,'','border:1px solid #000'),
    `<div class="hl" style="left:0;top:${ROWY[4]}px;width:${SHEET_W}px"></div>`,
    cell(0,6,0,6,'<b>SECCIÓN:</b>'), cell(1,6,3,6,esc(seccionDe(a)),'box','font-size:12px'),
    cell(5,6,6,6,'<b>PLANTA:</b>'), cell(7,6,8,6,'<b>1</b>','','justify-content:center'), cell(9,6,9,6,`<b>${+a.planta===1?'X':'--'}</b>`,'box'),
    cell(0,8,0,8,'<b>ANEXO:</b>'), cell(1,8,3,8,esc(a.anexo||'-'),'box'),
    cell(7,8,8,8,'<b>2</b>','','justify-content:center'), cell(9,8,9,8,`<b>${+a.planta!==1?'X':'--'}</b>`,'box'),
    `<div class="hl" style="left:0;top:${ROWY[9]}px;width:${SHEET_W}px"></div>`,
    cell(0,11,1,11,'<b>OBSERVACIÓN</b>'), cell(2,11,2,11,'<b>--</b>','box'), cell(4,11,5,11,'No conformidad'), cell(6,11,6,11,'<b>--</b>','box'),
    cell(0,13,1,13,'<b>LEVANTAMIENTO</b>'), cell(2,13,2,13,'<b>--</b>','box'), cell(4,13,5,13,'otros (especificar)'), cell(6,13,6,13,'<b>X</b>','box'),
    cell(8,13,10,13,'<span style="width:100%;text-align:center;border-bottom:1px solid #000">Acción del  cc</span>'),
    `<div class="hl" style="left:0;top:${ROWY[14]}px;width:${SHEET_W}px"></div><div class="hl" style="left:0;top:${ROWY[15]-2}px;width:${SHEET_W}px"></div>`,
    cell(0,16,11,17,esc(accionTexto(a)),'act'),
  ].join('');
  const frameL = `<div class="frame" style="left:0;top:${ROWY[3]}px;width:${SHEET_W}px;height:${ROWY[17]-ROWY[3]+blockH}px;border-width:1px"></div>`;
  const urls = await Promise.all(imgs.map(im => thumbFor(im)));
  const urlOf = im => urls[imgs.indexOf(im)];
  const pagesHtml = L.pages.map((pg, p) => {
    const imgsHtml = pg.map(it => `<img src="${urlOf(it.im)}" alt="" style="left:${it.x}px;top:${ROWY[17]+it.y}px;width:${it.w}px;height:${it.h}px">`).join('');
    const emptyMsg = !imgs.length ? `<div class="c" style="left:0;top:${ROWY[17]}px;width:${SHEET_W}px;height:${blockH}px;justify-content:center;color:#999;font-size:13px">Las imágenes pegadas aparecerán aquí, centradas</div>` : '';
    const sheet = `<div class="sheet" style="left:${LEFT}px;top:${TOP}px;width:${SHEET_W}px;transform:scale(${SCALE})">${frameL}${head}${imgsHtml}${emptyMsg}</div>`;
    return `<div class="pv-label">Hoja ${p+1} de ${L.pages.length}</div><div class="paper-box" style="aspect-ratio:794/1123"><div class="paper"><div class="hdr"><img src="data:image/jpeg;base64,${window.FASC_LOGO_B64}" alt="">GARANTÍA DE LA CALIDAD</div>${sheet}<div class="ftr">FASC-205-03</div></div></div>`;
  }).join('');
  const pv = $('#preview');
  pv.innerHTML = pagesHtml + `<p class="hint" style="align-self:flex-start;margin:0">Diseño: ${L.per} imagen${L.per>1?'es':''} por hoja · A4 vertical, escala 95 %, como la plantilla.</p>`;
  fitPaper();
}
function fitPaper(){ document.querySelectorAll('.paper-box').forEach(b => { const p=b.querySelector('.paper'); p.style.transform = `scale(${b.clientWidth/794})`; }); }
addEventListener('resize', debounce(fitPaper, 100));

/* ---------- Miniaturas ---------- */
async function thumbFor(im){
  if (thumbURL.has(im.storage_path)) return thumbURL.get(im.storage_path);
  const b = await DB.blob(im.storage_path); const u = URL.createObjectURL(b); thumbURL.set(im.storage_path, u); return u;
}
async function renderThumbs(){
  const list = evid.get(current.id) || [];
  const box = $('#thumbs');
  if (!list.length){ box.innerHTML = ''; return; }
  const urls = await Promise.all(list.map(thumbFor));
  box.innerHTML = list.map((im,i)=>`<div class="thumb" data-i="${i}">
     <div class="im" data-act="edit" title="Editar imagen"><img src="${urls[i]}" alt="Evidencia ${i+1}"></div>
     <div class="tb"><span class="mono">${i+1} · ${im.ancho}×${im.alto}</span><span style="display:flex">
       <button type="button" class="ib" data-act="up" title="Mover antes" ${i===0?'disabled':''}><svg viewBox="0 0 24 24"><path d="m15 18-6-6 6-6"/></svg></button>
       <button type="button" class="ib" data-act="down" title="Mover después" ${i===list.length-1?'disabled':''}><svg viewBox="0 0 24 24"><path d="m9 18 6-6-6-6"/></svg></button>
       <button type="button" class="ib" data-act="edit" title="Editar"><svg viewBox="0 0 24 24"><path d="M12 20h9"/><path d="M16.5 3.5a2.1 2.1 0 0 1 3 3L7 19l-4 1 1-4z"/></svg></button>
       <button type="button" class="ib danger" data-act="del" title="Eliminar"><svg viewBox="0 0 24 24"><path d="M3 6h18M8 6V4h8v2M6 6l1 14h10l1-14"/></svg></button>
     </span></div></div>`).join('');
}

/* ---------- Ingreso de imágenes ---------- */
function loadImage(blob){ return new Promise((res, rej) => { const u=URL.createObjectURL(blob); const i=new Image(); i.onload=()=>{ res(i); }; i.onerror=()=>rej(new Error('Imagen no válida')); i.src=u; }); }
async function prepImage(file){
  const img = await loadImage(file);
  let w = img.naturalWidth, h = img.naturalHeight; const MAX = 2400;
  const ok = /image\/(png|jpeg)/.test(file.type) && Math.max(w,h) <= MAX && file.size < 5e6;
  if (ok) return {blob:file, w, h};
  const s = Math.min(1, MAX/Math.max(w,h)); w=Math.round(w*s); h=Math.round(h*s);
  const cv = document.createElement('canvas'); cv.width=w; cv.height=h; const cx=cv.getContext('2d');
  cx.fillStyle='#fff'; cx.fillRect(0,0,w,h); cx.drawImage(img,0,0,w,h);
  const png = file.type==='image/png' || file.type==='image/gif' || file.type==='image/bmp';
  const blob = await new Promise(r=>cv.toBlob(r, png?'image/png':'image/jpeg', .92));
  return {blob, w, h};
}
const extOf = b => /png/.test(b.type) ? 'png' : 'jpg';
async function addImages(files){
  files = [...files].filter(f => f && /^image\//.test(f.type));
  if (!files.length || !current) return;
  const a = current; const busy = $('#busy'); busy.hidden=false;
  try{
    for (let i=0;i<files.length;i++){
      busy.textContent = `Guardando imagen ${i+1} de ${files.length}…`;
      const {blob, w, h} = await prepImage(files[i]);
      const path = `${a.codigo_cc.replace(/[^\w-]+/g,'_')}/${a.num_accion}/${uid()}.${extOf(blob)}`;
      await DB.upload(path, blob);
      const list = evid.get(a.id) || []; evid.set(a.id, list);
      const row = await DB.insertEvid({accion_id:a.id, storage_path:path, original_path:path, orden:list.length ? list[list.length-1].orden+1 : 0, ancho:w, alto:h, mime:blob.type});
      list.push(row);
    }
    toast(files.length===1 ? 'Imagen guardada' : `${files.length} imágenes guardadas`);
  } catch(err){ console.error(err); toast('No se pudo guardar la imagen: '+(err.message||err), 5000); }
  busy.hidden = true;
  refreshDrawer(); renderTable();
}
async function moveImage(i, d){
  const list = evid.get(current.id); const j=i+d; if (j<0||j>=list.length) return;
  [list[i], list[j]] = [list[j], list[i]];
  list.forEach((im,k)=>im.orden=k);
  refreshDrawer();
  try{ await Promise.all(list.map(im=>DB.updateEvid(im.id,{orden:im.orden}))); } catch(err){ toast('No se pudo guardar el orden'); }
}
async function deleteImage(i, btn){
  if (btn.dataset.confirm !== '1'){ btn.dataset.confirm='1'; btn.title='Haz clic otra vez para eliminar'; btn.style.background='var(--accent-weak)'; toast('Haz clic otra vez en el ícono para eliminar la imagen'); setTimeout(()=>{ btn.dataset.confirm=''; btn.style.background=''; }, 3000); return; }
  const list = evid.get(current.id); const im = list[i];
  try{
    await DB.deleteEvid(im.id);
    await DB.removeFiles([im.storage_path, im.original_path!==im.storage_path ? im.original_path : null]);
    list.splice(i,1); toast('Imagen eliminada'); refreshDrawer(); renderTable();
  } catch(err){ toast('No se pudo eliminar: '+(err.message||err), 5000); }
}

/* ============================================================
   Editor de imágenes: rotar, recortar, resaltar (rectángulo) y flecha
   ============================================================ */
const Editor = (() => {
  const el = $('#editor'), cv = $('#edCanvas'), cx = cv.getContext('2d');
  let work = null, stack = [], tool = null, drag = null, target = null, color = '#e00000';
  const COLORS = ['#e00000','#ffcc00','#0070c0','#00a651'];
  $('#edColors').innerHTML = COLORS.map(c=>`<button type="button" class="color-dot" data-color="${c}" style="background:${c}" aria-label="Color ${c}"></button>`).join('');
  const mk = (w,h) => { const c=document.createElement('canvas'); c.width=w; c.height=h; return c; };
  const clone = src => { const c=mk(src.width,src.height); c.getContext('2d').drawImage(src,0,0); return c; };
  function paint(){
    if (cv.width!==work.width || cv.height!==work.height){ cv.width = work.width; cv.height = work.height; }
    fit(); cx.clearRect(0,0,cv.width,cv.height); cx.drawImage(work,0,0);
    if (drag && drag.x2!=null){
      const r = rectOf(drag);
      if (tool==='crop'){ cx.fillStyle='rgba(0,0,0,.55)'; cx.fillRect(0,0,cv.width,cv.height); cx.drawImage(work, r.x,r.y,r.w,r.h, r.x,r.y,r.w,r.h); cx.strokeStyle='#fff'; cx.lineWidth=Math.max(1,lw()/2); cx.setLineDash([lw()*2,lw()*1.5]); cx.strokeRect(r.x,r.y,r.w,r.h); cx.setLineDash([]); }
      else drawShape(cx, drag);
    }
  }
  function fit(){ const st=$('#edStage'); const W=st.clientWidth-32, H=st.clientHeight-32; const k=Math.min(W/work.width, H/work.height, 1); cv.style.width=Math.floor(work.width*k)+'px'; cv.style.height=Math.floor(work.height*k)+'px'; }
  addEventListener('resize', () => { if (work && !el.hidden) fit(); });
  const lw = () => Math.max(3, Math.round(Math.max(work.width, work.height)/250));
  const rectOf = d => ({x:Math.min(d.x1,d.x2), y:Math.min(d.y1,d.y2), w:Math.abs(d.x2-d.x1), h:Math.abs(d.y2-d.y1)});
  function drawShape(c, d){
    c.strokeStyle = color; c.fillStyle = color; c.lineWidth = lw(); c.lineJoin='round'; c.lineCap='round';
    if (tool==='rect'){ const r=rectOf(d); c.strokeRect(r.x,r.y,r.w,r.h); }
    if (tool==='arrow'){ const ang=Math.atan2(d.y2-d.y1, d.x2-d.x1), hl=lw()*4.5;
      c.beginPath(); c.moveTo(d.x1,d.y1); c.lineTo(d.x2-Math.cos(ang)*hl*.6, d.y2-Math.sin(ang)*hl*.6); c.stroke();
      c.beginPath(); c.moveTo(d.x2,d.y2); c.lineTo(d.x2-hl*Math.cos(ang-.45), d.y2-hl*Math.sin(ang-.45)); c.lineTo(d.x2-hl*Math.cos(ang+.45), d.y2-hl*Math.sin(ang+.45)); c.closePath(); c.fill(); }
  }
  const pos = e => { const r=cv.getBoundingClientRect(); return {x:(e.clientX-r.left)*cv.width/r.width, y:(e.clientY-r.top)*cv.height/r.height}; };
  cv.addEventListener('pointerdown', e => { if(!tool) return; cv.setPointerCapture(e.pointerId); const p=pos(e); drag={x1:p.x,y1:p.y,x2:null,y2:null}; });
  cv.addEventListener('pointermove', e => { if(!drag) return; const p=pos(e); drag.x2=Math.max(0,Math.min(cv.width,p.x)); drag.y2=Math.max(0,Math.min(cv.height,p.y)); paint(); });
  cv.addEventListener('pointerup', () => {
    if (!drag || drag.x2==null){ drag=null; return; }
    const r = rectOf(drag);
    if (r.w>4 || r.h>4){
      stack.push(clone(work));
      if (tool==='crop'){ if (r.w>8 && r.h>8){ const n=mk(Math.round(r.w),Math.round(r.h)); n.getContext('2d').drawImage(work, r.x,r.y,r.w,r.h, 0,0,n.width,n.height); work=n; } else stack.pop(); setTool(null); }
      else drawShape(work.getContext('2d'), drag);
    }
    drag=null; paint();
  });
  function rotate(dir){ stack.push(clone(work)); const n=mk(work.height, work.width), c=n.getContext('2d'); c.translate(n.width/2,n.height/2); c.rotate(dir*Math.PI/2); c.drawImage(work,-work.width/2,-work.height/2); work=n; paint(); }
  function setTool(t){ tool = tool===t ? null : t; el.querySelectorAll('[data-tool]').forEach(b=>b.classList.toggle('on', b.dataset.tool===tool));
    $('#edTip').textContent = tool==='crop' ? 'Arrastra para marcar el área que quieres conservar.' : tool==='rect' ? 'Arrastra para dibujar un recuadro de resaltado.' : tool==='arrow' ? 'Arrastra desde el inicio hasta la punta de la flecha.' : 'Elige una herramienta y arrastra sobre la imagen.';
    cv.style.cursor = tool ? 'crosshair' : 'default'; }
  async function fromPath(path){ const img = await loadImage(await DB.blob(path)); const c=mk(img.naturalWidth,img.naturalHeight); c.getContext('2d').drawImage(img,0,0); return c; }
  async function open(im){ target = im; stack=[]; setTool(null); work = await fromPath(im.storage_path); el.hidden=false; paint(); }
  function close(){ el.hidden=true; target=null; work=null; }
  async function save(){
    if (!stack.length){ close(); return; }
    const im = target, png = !/jpe?g/.test(im.mime || '');
    const blob = await new Promise(r=>work.toBlob(r, png?'image/png':'image/jpeg', .92));
    const a = acciones.find(x=>x.id===im.accion_id);
    const path = `${a.codigo_cc.replace(/[^\w-]+/g,'_')}/${a.num_accion}/${uid()}.${extOf(blob)}`;
    try{
      await DB.upload(path, blob);
      const old = im.storage_path;
      await DB.updateEvid(im.id, {storage_path:path, ancho:work.width, alto:work.height, mime:blob.type, original_path: im.original_path || old});
      if (old !== (im.original_path || old)) await DB.removeFiles([old]);
      im.original_path = im.original_path || old; im.storage_path = path; im.ancho = work.width; im.alto = work.height; im.mime = blob.type;
      close(); toast('Imagen actualizada'); refreshDrawer();
    } catch(err){ toast('No se pudo guardar la edición: '+(err.message||err), 5000); }
  }
  async function restore(){ if (!target?.original_path) return; stack.push(clone(work)); work = await fromPath(target.original_path); paint(); toast('Se cargó la imagen original. Guarda para confirmar.'); }
  el.addEventListener('click', e => {
    const b = e.target.closest('button'); if (!b) return;
    if (b.dataset.tool) setTool(b.dataset.tool);
    if (b.dataset.color){ color=b.dataset.color; el.querySelectorAll('[data-color]').forEach(x=>x.style.outline = x===b ? '2px solid #fff' : ''); }
    const a = b.dataset.ed;
    if (a==='rotl') rotate(-1); if (a==='rotr') rotate(1);
    if (a==='undo' && stack.length){ work = stack.pop(); paint(); }
    if (a==='orig') restore(); if (a==='cancel') close(); if (a==='save') save();
  });
  return {open, close};
})();

/* ---------- Panel lateral ---------- */
function openDrawer(id){
  current = acciones.find(a=>a.id===id); if (!current) return;
  $('#scrim').hidden=false; $('#drawer').hidden=false; document.body.style.overflow='hidden';
  refreshDrawer(true); renderTable();
}
function closeDrawer(){ $('#scrim').hidden=true; $('#drawer').hidden=true; document.body.style.overflow=''; current=null; renderTable(); }
function refreshDrawer(fillForm){
  const a = current; if (!a) return;
  $('#dTitle').innerHTML = `<span class="mono">${esc(a.codigo_cc)}</span> · Acción N° ${esc(a.num_accion)}`;
  $('#dSub').textContent = a.activo===false ? 'Esta acción ya no figura en el último Excel importado.' : (a.generado_at ? 'Formato descargado el '+new Date(a.generado_at).toLocaleString('es-PE') : 'Formato aún no descargado');
  $('#meta').innerHTML = [['Descripción', a.descripcion],['Responsable', a.responsable],['Área', a.area],['Fecha', fmtDate(a.fecha)],['Status de acciones', a.status_acciones],['Status', a.status],['Comentario', a.comentario]]
    .map(([k,v])=>`<dt>${k}</dt><dd>${esc(v||'—').replace(/\n/g,'<br>')}</dd>`).join('');
  if (fillForm){
    $('#fSeccion').value = seccionDe(a); $('#fAnexo').value = a.anexo || '-'; $('#fTexto').value = accionTexto(a); $('#fPorPag').value = a.por_pagina || '';
  }
  document.querySelectorAll('#fPlanta button').forEach(b=>b.classList.toggle('on', +b.dataset.v === (+a.planta||2)));
  $('#fname').textContent = fileName(a);
  renderThumbs(); renderPreview();
}
const saveField = debounce(async (id, patch) => { try{ await DB.updateAccion(id, patch); } catch(err){ toast('No se pudo guardar el cambio: '+(err.message||err), 5000); } }, 500);
function setField(k, v){ if (!current) return; current[k] = v; saveField(current.id, {seccion:current.seccion, anexo:current.anexo, texto_accion:current.texto_accion, planta:current.planta, por_pagina:current.por_pagina}); renderPreview(); $('#fname').textContent = fileName(current); }
$('#fSeccion').addEventListener('input', e => setField('seccion', e.target.value));
$('#fAnexo').addEventListener('input', e => setField('anexo', e.target.value || '-'));
$('#fTexto').addEventListener('input', e => setField('texto_accion', e.target.value));
$('#fPorPag').addEventListener('change', e => setField('por_pagina', e.target.value ? +e.target.value : null));
$('#fPlanta').addEventListener('click', e => { const b=e.target.closest('button'); if(!b) return; setField('planta', +b.dataset.v); document.querySelectorAll('#fPlanta button').forEach(x=>x.classList.toggle('on', x===b)); });
$('#btnResetTexto').addEventListener('click', () => { setField('texto_accion', null); $('#fTexto').value = accionTexto(current); });

$('#btnDownload').addEventListener('click', async () => {
  const a = current; if (!a) return; const btn = $('#btnDownload'); btn.disabled = true;
  try{
    const imgs = evid.get(a.id) || [];
    if (!imgs.length) toast('Se descargará el formato sin imágenes de evidencia.');
    const blob = await buildXlsx(a, imgs);
    const u = URL.createObjectURL(blob), link = document.createElement('a');
    link.href = u; link.download = fileName(a); document.body.appendChild(link); link.click(); link.remove();
    setTimeout(()=>URL.revokeObjectURL(u), 4000);
    a.generado_at = new Date().toISOString();
    await DB.updateAccion(a.id, {generado_at:a.generado_at});
    refreshDrawer(); renderTable();
  } catch(err){ console.error(err); toast('No se pudo generar el archivo: '+(err.message||err), 6000); }
  btn.disabled = false;
});

/* ---------- Eventos ---------- */
$('#thead').addEventListener('click', e => { const b=e.target.closest('.fbtn'); if (b) openFilter(b.dataset.col, b); });
$('#tbody').addEventListener('click', e => { const tr=e.target.closest('tr[data-id]'); if (tr) openDrawer(tr.dataset.id); });
$('#chips').addEventListener('click', e => { const b=e.target.closest('[data-clear]'); if (b){ delete filters[b.dataset.clear]; renderTable(); } });
$('#btnClear').addEventListener('click', () => { filters={}; sort={k:null,dir:1}; query=''; $('#q').value=''; renderTable(); });
$('#q').addEventListener('input', debounce(e => { query=e.target.value.trim(); renderTable(); }, 150));
$('#showInactive').addEventListener('change', renderTable);
$('#btnImport').addEventListener('click', () => $('#fileImport').click());
$('#fileImport').addEventListener('change', e => { const f=e.target.files[0]; e.target.value=''; if (f) importFile(f); });
$('#btnClose').addEventListener('click', closeDrawer);
$('#scrim').addEventListener('click', closeDrawer);
const drop = $('#drop');
drop.addEventListener('click', () => $('#fileImg').click());
drop.addEventListener('keydown', e => { if (e.key==='Enter' || e.key===' '){ e.preventDefault(); $('#fileImg').click(); } });
$('#fileImg').addEventListener('change', e => { addImages(e.target.files); e.target.value=''; });
['dragenter','dragover'].forEach(t=>drop.addEventListener(t, e=>{ e.preventDefault(); drop.classList.add('over'); }));
['dragleave','drop'].forEach(t=>drop.addEventListener(t, e=>{ e.preventDefault(); drop.classList.remove('over'); }));
drop.addEventListener('drop', e => addImages(e.dataTransfer.files));
document.addEventListener('paste', e => {
  if ($('#drawer').hidden || !$('#editor').hidden) return;
  const files = [...(e.clipboardData?.items||[])].filter(i=>i.kind==='file' && i.type.startsWith('image/')).map(i=>i.getAsFile());
  if (files.length){ e.preventDefault(); addImages(files); }
});
$('#thumbs').addEventListener('click', e => {
  const b = e.target.closest('[data-act]'); if (!b) return; const i = +b.closest('.thumb').dataset.i; const list = evid.get(current.id);
  if (b.dataset.act==='up') moveImage(i,-1); if (b.dataset.act==='down') moveImage(i,1);
  if (b.dataset.act==='edit') Editor.open(list[i]); if (b.dataset.act==='del') deleteImage(i, b);
});

/* ---------- Inicio ---------- */
(async () => {
  const m = $('#mode');
  if (sb){ m.textContent='Supabase conectado'; m.className='mode cloud'; } else { m.textContent='Modo local'; m.className='mode local'; notice('Modo local: no hay Supabase configurado en config.js, los cambios se pierden al cerrar la página.'); }
  try { await reload(); }
  catch(err){ console.error(err); m.textContent='Error de conexión'; m.className='mode local';
    notice(/cc_acciones|schema cache|does not exist/.test(err.message||'') ? 'Supabase está conectado, pero faltan las tablas. Ejecuta supabase/setup.sql en el SQL Editor de tu proyecto y recarga la página.' : 'No se pudo leer Supabase: '+(err.message||err), 'err');
    renderTable(); }
  window.__app = {buildXlsx, computeLayout, parseCC, get acciones(){return acciones;}, get evid(){return evid;}};
})();
})();
