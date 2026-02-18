/* global Office, Word */
let exigencesCache = [];
let mappings = {}; // { clauseId: { number: '6.2.3', exig: ['1.1.2(a)', ...] } }
let hazards = [];   // [{ id:'HZ-001', label:'...', category:'...', clauseIds:['C-xxxxx'] }]
let config = {};

Office.onReady(async () => {
  await loadConfig();
  await loadExigences();
  bindTabs();
  document.getElementById('refresh').onclick = refreshFromDocument;
  document.getElementById('saveMapping').onclick = saveMappingLocal;
  document.getElementById('genZA').onclick = generateZA;
  document.getElementById('addHazard').onclick = addHazard;
  document.getElementById('syncSP').onclick = () => syncCoverageToSP(false);
  document.getElementById('syncHazardsSP').onclick = () => syncHazardsToSP(false);
  await refreshFromDocument();
});

function bindTabs(){
  document.querySelectorAll('.tab-btn').forEach(btn=>{
    btn.onclick = ()=>{
      document.querySelectorAll('.tab').forEach(t=>t.classList.remove('active'));
      const tab = btn.dataset.tab;
      document.getElementById('tab-'+tab).classList.add('active');
    };
  });
}

async function loadConfig(){
  const res = await fetch('config.json');
  config = await res.json();
  document.getElementById('cfg').innerText = JSON.stringify(config, null, 2);
}

async function loadExigences(){
  const res = await fetch('assets/exigences.json');
  exigencesCache = await res.json();
  const list = document.getElementById('exigences-list');
  list.innerHTML = '';
  exigencesCache.forEach(ex => {
    const div = document.createElement('label');
    div.className = 'ex-item';
    div.innerHTML = `<input type=\"checkbox\" value=\"${ex.id}\"/> <strong>${ex.id}</strong> – ${ex.description}`;
    list.appendChild(div);
  });
}

async function refreshFromDocument(){
  await Word.run(async context => {
    const sel = context.document.getSelection();
    sel.load('paragraphs');
    await context.sync();
    if(sel.paragraphs.items.length === 0){ return; }
    const p = sel.paragraphs.items[0];

    const text = p.text.trim();
    const number = text.split(' ')[0];
    document.getElementById('clause-number').innerText = number || '--';

    const id = await ensureContentControlId(p, context);
    document.getElementById('clause-id').innerText = id;

    await loadJsonCc(context);

    const current = mappings[id];
    document.querySelectorAll('#exigences-list input[type=checkbox]').forEach(cb => cb.checked = false);
    if(current && current.exig){
      const set = new Set(current.exig);
      document.querySelectorAll('#exigences-list input[type=checkbox]').forEach(cb => {
        cb.checked = set.has(cb.value);
      });
    }
  });
}

async function ensureContentControlId(p, context){
  const ccs = p.contentControls; ccs.load('items'); await context.sync();
  if(ccs.items.length>0) return ccs.items[0].tag || ccs.items[0].id.toString();
  const newId = 'C-' + Math.floor(10000 + Math.random()*90000);
  const cc = p.insertContentControl(); cc.tag = newId; cc.appearance = 'Hidden';
  await context.sync(); return newId;
}

async function loadJsonCc(context){
  const body = context.document.body; const ccs = body.contentControls; ccs.load('items'); await context.sync();
  let jsonCc = null; for(const cc of ccs.items){ if(cc.tag==='ANNEXE_Z_JSON'){ jsonCc = cc; break; } }
  if(!jsonCc){ jsonCc = body.insertContentControl(); jsonCc.tag='ANNEXE_Z_JSON'; jsonCc.appearance='Hidden'; jsonCc.insertText('{}','Replace'); }
  jsonCc.load('text'); await context.sync();
  try { const store = JSON.parse(jsonCc.text||'{}'); mappings = store.mappings||{}; hazards = store.hazards||[]; }
  catch(e){ mappings={}; hazards=[]; }
  renderHazards();
}

async function saveJsonCc(context){
  const body = context.document.body; const ccs = body.contentControls; ccs.load('items'); await context.sync();
  let jsonCc = null; for(const cc of ccs.items){ if(cc.tag==='ANNEXE_Z_JSON'){ jsonCc = cc; break; } }
  if(!jsonCc){ jsonCc = body.insertContentControl(); jsonCc.tag='ANNEXE_Z_JSON'; jsonCc.appearance='Hidden'; }
  jsonCc.insertText(JSON.stringify({mappings, hazards}),'Replace');
  await context.sync();
}

async function saveMappingLocal(){
  await Word.run(async context => {
    const id = document.getElementById('clause-id').innerText;
    const number = document.getElementById('clause-number').innerText;
    const selected = [...document.querySelectorAll('#exigences-list input:checked')].map(x=>x.value);
    if(!id || id==='--') return;
    mappings[id] = { number, exig: selected };
    await saveJsonCc(context);
  });
  alert('Lien enregistré dans le document.');
}

function addHazard(){
  const label = document.getElementById('hzLabel').value.trim();
  const category = document.getElementById('hzCategory').value.trim();
  if(!label) return;
  const hz = { id: 'HZ-'+String(1000+hazards.length), label, category, clauseIds: [] };
  hazards.push(hz); document.getElementById('hzLabel').value=''; document.getElementById('hzCategory').value='';
  renderHazards();
}

function renderHazards(){
  const wrap = document.getElementById('hazards-list');
  if(hazards.length===0){ wrap.innerHTML='[Aucun pour l’instant]'; return; }
  wrap.innerHTML='';
  hazards.forEach(hz=>{
    const div = document.createElement('div'); div.className='hz';
    div.innerHTML = `<strong>${hz.id}</strong> — ${hz.label} (${hz.category||''})`;
    wrap.appendChild(div);
  });
}

async function generateZA(){
  await Word.run(async context => {
    const reverse = {};
    Object.values(mappings).forEach(m => { (m.exig||[]).forEach(e => { (reverse[e]=reverse[e]||new Set()).add(m.number); }); });
    const doc = context.document; const search = doc.body.search('[ANNEXE_Z_TABLE]', { matchWildcards:false }); search.load('items'); await context.sync();
    let range; if(search.items.length){ range = search.items[0]; range.insertText('', 'Replace'); } else { range = doc.body.paragraphs.getLast().getRange('End'); }
    range.insertParagraph('Table ZA.1 — Correspondence between this European Standard and Annex III of Regulation (EU) 2023/1230', 'Heading 2');
    const rows = [];
    rows.push(['The relevant Essential Requirements of Regulation (EU) 2023/1230','Clause(s)/sub-clause(s) of this EN','Remarks/Notes']);
    exigencesCache.forEach(ex => {
      const clauses = reverse[ex.id] ? Array.from(reverse[ex.id]).sort((a,b)=>a.localeCompare(b,undefined,{numeric:true})) : [];
      rows.push([`${ex.id} — ${ex.description}`, clauses.join(', '), '']);
    });
    range.insertTable(rows.length, 3, 'Start');
    const tbl = range.paragraphs.getFirst().previous(); await context.sync();
    for(let r=0; r<rows.length; r++){ for(let c=0; c<3; c++){ tbl.getCell(r,c).insertParagraph(rows[r][c], 'Normal'); } }
    await context.sync(); alert('Annexe Z mise à jour.');
  });
}

async function syncCoverageToSP(dryRun){ return window.SP_SYNC.syncCoverage(mappings, config, dryRun); }
async function syncHazardsToSP(dryRun){ return window.SP_SYNC.syncHazards(hazards, config, dryRun); }
