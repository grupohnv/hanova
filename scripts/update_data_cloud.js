const fs = require('fs');

const htmlPath = './dashboard_hanova.html';
const b64Path  = '/tmp/hanova/hanova_data.b64';
const jsonPath = '/tmp/hanova/hanova_data.json';

let html = fs.readFileSync(htmlPath, 'utf8');
const b64  = fs.readFileSync(b64Path, 'utf8').trim();
const meta = JSON.parse(fs.readFileSync(jsonPath, 'utf8')).meta;

const geradoEm = new Date(meta.geradoEm).toLocaleDateString('pt-BR', {
  day:'2-digit', month:'2-digit', year:'numeric', hour:'2-digit', minute:'2-digit'
});

const blockStart = /\/\* ═+\s*EMBEDDED DATA[\s\S]*?document\.addEventListener\('DOMContentLoaded', loadEmbeddedData\);\s*/g;
const cleaned = html.replace(blockStart, '');
const removedCount = (html.match(blockStart)||[]).length;
console.log('Removed ' + removedCount + ' existing data block(s)');

const newBlock = `/* ══════════════════════════════════════════════════
   EMBEDDED DATA (auto-load, sem upload)
   Gerado em: ${geradoEm}
══════════════════════════════════════════════════ */
const EMBEDDED_B64GZ = '${b64}';
const DADOS_ATUALIZADOS_EM = '${geradoEm}';

async function loadEmbeddedData() {
  document.getElementById('loadingOverlay').style.display = 'flex';
  try {
    const binStr = atob(EMBEDDED_B64GZ);
    const bytes = new Uint8Array(binStr.length);
    for (let i = 0; i < binStr.length; i++) bytes[i] = binStr.charCodeAt(i);
    const ds = new DecompressionStream('gzip');
    const writer = ds.writable.getWriter();
    writer.write(bytes); writer.close();
    const chunks = []; const reader = ds.readable.getReader();
    while (true) { const {done,value} = await reader.read(); if(done) break; chunks.push(value); }
    const total = chunks.reduce((s,c)=>s+c.length,0);
    const buf = new Uint8Array(total); let off=0;
    for (const c of chunks) { buf.set(c,off); off+=c.length; }
    const data = JSON.parse(new TextDecoder().decode(buf));

    cadastroMap.clear();
    data.cadastro.forEach(([cod,nome,vend,loja,codRede,tipo])=>{ if(cod) cadastroMap.set(String(cod),{nomecli:nome||'',nomevend:vend||'',loja:loja||'',cod_rede:codRede||'',tipo:(tipo||'').toUpperCase()}); });
    produtosAtivos.clear();
    data.produtos.forEach(p=>{ if(p) produtosAtivos.add(String(p).trim().toUpperCase()); });
    metaMap.clear();
    (data.metas||[]).forEach(function(r){ const c=String(r[0]||'').trim(),v=typeof r[1]==='number'?r[1]:0; if(c&&v>0) metaMap.set(c,(metaMap.get(c)||0)+v); });
    dadosBase = data.dados.map(([codcli,valor,anomes,descprod])=>{
      const key=String(codcli||'');
      const cad=cadastroMap.get(key)||{nomecli:'',nomevend:'N/D'};
      const ano=Math.floor(anomes/100), mes=anomes%100;
      return {codcli:key, nomecli:cad.nomecli, nomevend:cad.nomevend,
        valor:typeof valor==='number'?valor:0,
        _data:new Date(ano,mes,0), _anomes:anomes,
        descprod:String(descprod||'').trim().toUpperCase()};
    });
    bonifBase = (data.bonificacoes||[]).map(([codcli,valor,anomes])=>{
      const key=String(codcli||'');
      const cad=cadastroMap.get(key)||{nomecli:'',nomevend:'N/D'};
      const ano=Math.floor(anomes/100), mes=anomes%100;
      return {codcli:key, nomecli:cad.nomecli, nomevend:cad.nomevend,
        valor:typeof valor==='number'?valor:0,
        _data:new Date(ano,mes,0), _anomes:anomes};
    });
    buildFilters();
    document.getElementById('uploadScreen').style.display='none';
    document.getElementById('loadingOverlay').style.display='none';
    document.getElementById('dashboard').style.display='block';
    const el=document.getElementById('dadosAtualizadosEm');
    if(el) el.textContent='Dados de '+DADOS_ATUALIZADOS_EM;
    applyFilters();
  } catch(e) {
    document.getElementById('loadingOverlay').style.display='none';
    const er=document.getElementById('loadingError');
    if(er){er.style.display='block';er.textContent='Erro: '+e.message;}
    console.error(e);
  }
}

document.addEventListener('DOMContentLoaded', loadEmbeddedData);
`;

if (!cleaned.includes('// Drag-and-drop support')) {
  console.error('ERROR: anchor not found!');
  process.exit(1);
}
const out = cleaned.replace('// Drag-and-drop support', newBlock + '\n// Drag-and-drop support');
fs.writeFileSync(htmlPath, out, 'utf8');
const size = fs.statSync(htmlPath).size;
console.log('Done! HTML: ' + (size/1024/1024).toFixed(2) + ' MB | Data: ' + geradoEm);
