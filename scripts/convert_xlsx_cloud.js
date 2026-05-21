const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

function normCol(headers, candidates) {
  return headers.find(h => candidates.some(c => h.trim().toLowerCase().includes(c.toLowerCase())));
}

function parseDate(val) {
  if (!val) return null;
  if (val instanceof Date) return val;
  if (typeof val === 'number') {
    const d = XLSX.SSF.parse_date_code(val);
    if (d) return new Date(d.y, d.m - 1, d.d);
  }
  if (typeof val === 'string') {
    const s = val.trim();
    const m1 = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    if (m1) return new Date(+m1[3], +m1[2]-1, +m1[1]);
    const m2 = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (m2) return new Date(+m2[1], +m2[2]-1, +m2[3]);
  }
  return null;
}

function parseValor(val) {
  if (typeof val === 'number') return val;
  if (!val) return 0;
  const s = String(val).replace(/[R$\s]/g, '').trim();
  const hasDot = s.includes('.');
  const hasComma = s.includes(',');
  if (hasDot && hasComma) {
    const lastDot = s.lastIndexOf('.');
    const lastComma = s.lastIndexOf(',');
    if (lastComma > lastDot) return parseFloat(s.replace(/\./g, '').replace(',', '.')) || 0;
    else return parseFloat(s.replace(/,/g, '')) || 0;
  }
  if (hasComma && !hasDot) return parseFloat(s.replace(',', '.')) || 0;
  return parseFloat(s) || 0;
}

// ---- DADOS PROCESSADOS ----
console.log('Reading dados_processados...');
const wbDados = XLSX.readFile('/tmp/hanova/dados_relatorios/dados_processados_powerbi_novo.xlsx', { cellDates: true, raw: true });
const wsDados = wbDados.Sheets[wbDados.SheetNames[0]];
const rawDados = XLSX.utils.sheet_to_json(wsDados, { defval: null, raw: true });
console.log('Rows:', rawDados.length);

const headers = rawDados.length > 0 ? Object.keys(rawDados[0]) : [];
console.log('Columns:', headers.slice(0, 10).join(', '));

const colCodcli = normCol(headers, ['codcli','cod_cli','codparc','codigo']);
const colValor  = normCol(headers, ['valor','vlr','vl_bruto','vl_liq']);
const colPeddata= normCol(headers, ['peddata','data','dtped','datpedido']);
const colDesc   = normCol(headers, ['descprod','desc_prod','produto','descricao','nomeprod']);
console.log('Cols found:', {colCodcli, colValor, colPeddata, colDesc});

const DATA_INICIAL = new Date('2025-01-01');
const dadosCompact = [];

for (const r of rawDados) {
  const dt = parseDate(r[colPeddata]);
  if (!dt || dt < DATA_INICIAL) continue;
  const valor = parseValor(r[colValor]);
  if (!valor && valor !== 0) continue;
  const anomes = dt.getFullYear() * 100 + (dt.getMonth() + 1);
  dadosCompact.push([
    String(r[colCodcli] || ''),
    valor,
    anomes,
    String(r[colDesc] || '')
  ]);
}
console.log('Filtered rows:', dadosCompact.length);

// ---- CADASTRO CLIENTE ----
console.log('Reading cadastro...');
const wbCad = XLSX.readFile('/tmp/hanova/dados_relatorios/cadastro-cliente-dimensão.xlsx', { cellDates: true, raw: true });
const wsCad = wbCad.Sheets[wbCad.SheetNames[0]];
const rawCad = XLSX.utils.sheet_to_json(wsCad, { defval: null, raw: true });
console.log('Rows:', rawCad.length);
const hCad = rawCad.length > 0 ? Object.keys(rawCad[0]) : [];
const cCod = normCol(hCad, ['codcli','cod_cli','codparc','codigo']);
const cNome= normCol(hCad, ['nomecli','nome_cli','nomeparc','nome']);
const cVend= normCol(hCad, ['nomevend','nome_vend','vendedor','representante']);
console.log('Cad cols:', {cCod, cNome, cVend});

const cadCompact = rawCad.map(r => [
  String(r[cCod] || ''),
  String(r[cNome] || ''),
  String(r[cVend] || '')
]);

// ---- PRODUTOS HANOVA ----
console.log('Reading produtos...');
const wbProd = XLSX.readFile('/tmp/hanova/dados_relatorios/Produtos Hanova.xlsx', { cellDates: true, raw: true });
const wsProd = wbProd.Sheets[wbProd.SheetNames[0]];
const rawProd = XLSX.utils.sheet_to_json(wsProd, { defval: null, raw: true });
console.log('Rows:', rawProd.length);
const hProd = rawProd.length > 0 ? Object.keys(rawProd[0]) : [];
console.log('Prod cols:', hProd.slice(0, 5));
const cProd = normCol(hProd, ['produto','descprod','desc_prod','descricao','nome','product','item']) || hProd[0];
console.log('Using prod col:', cProd);

const prodCompact = rawProd.map(r => String(r[cProd] || '')).filter(Boolean);

// ---- METAS POR CLIENTE ----
// Lê todas as planilhas de meta da pasta mais recente em G:/Meu Drive/comercial hanova/metas/
// Estrutura de cada arquivo: linha 0 = mapeamento de colunas, linhas 1+ = dados
// Coluna VALIDAR META (penúltima) = valor da meta por cliente
const metasCompact = [];  // [codcli, metaValor, nomevend, mes_ref]
try {
  const META_BASE = '/tmp/hanova/metas/';
  const subfolders = fs.readdirSync(META_BASE)
    .filter(f => /^\d{2}-\d{4}$/.test(f))
    .sort()
    .reverse(); // mais recente primeiro

  if (subfolders.length > 0) {
    const mesRef = subfolders[0]; // ex: "05-2026"
    const metaDir = path.join(META_BASE, mesRef);
    const metaFiles = fs.readdirSync(metaDir)
      .filter(f => f.endsWith('.xlsx') && !f.startsWith('~') && !f.startsWith('META_05_2026'));

    console.log(`Reading metas from ${mesRef} (${metaFiles.length} files)...`);
    let totalMeta = 0;

    for (const file of metaFiles) {
      try {
        const wbM = XLSX.readFile(path.join(metaDir, file), { raw: true });
        const wsM = wbM.Sheets[wbM.SheetNames[0]];
        const rowsM = XLSX.utils.sheet_to_json(wsM, { defval: null, raw: true });
        if (rowsM.length < 2) continue;

        // Row 0 = col name map (values are the real column names)
        const colMap = rowsM[0];
        const keys = Object.keys(colMap);
        const kCodcli   = keys.find(k => colMap[k] === 'CODCLI');
        const kValidar  = keys.find(k => colMap[k] === 'VALIDAR META');
        const kNomevend = keys.find(k => colMap[k] === 'NOMEVEND');
        const kMesRef   = mesRef; // "MM-YYYY"

        if (!kCodcli || !kValidar) continue;

        for (const r of rowsM.slice(1)) {
          const codcli = r[kCodcli];
          const metaVal = r[kValidar];
          if (!codcli || typeof metaVal !== 'number' || metaVal <= 0) continue;
          const nomevend = String(r[kNomevend] || '');
          metasCompact.push([String(codcli), metaVal, nomevend, kMesRef]);
          totalMeta++;
        }
      } catch(e) {
        console.warn('  Skip', file, e.message);
      }
    }
    console.log(`Metas loaded: ${totalMeta} clientes com meta em ${mesRef}`);
  } else {
    console.log('No meta folders found, skipping.');
  }
} catch(e) {
  console.warn('Meta read error:', e.message);
}

// ---- BONIFICAÇÕES ----
const bonifCompact = [];  // [codcli, valor, anomes]
try {
  const dadosDir = '/tmp/hanova/dados_relatorios/';
  const allFiles = fs.readdirSync(dadosDir);
  const bonifFiles = allFiles.filter(f =>
    f.toLowerCase().includes('bonificad') && (f.endsWith('.xlsx') || f.endsWith('.xls')) && !f.startsWith('~'));

  if (bonifFiles.length > 0) {
    // Usar o mais recente pelo nome (ordenação alfabética funciona p/ o padrão "08-24 A 07-25")
    const bonifFile = bonifFiles.sort().reverse()[0];
    console.log('Reading bonificacoes:', bonifFile);
    const wbBonif = XLSX.readFile(path.join(dadosDir, bonifFile), { cellDates: true, raw: true });

    // Tentar aba "new sheet", senão usar a primeira
    const sheetName = wbBonif.SheetNames.find(s =>
      s.toLowerCase().includes('new') || s.toLowerCase().includes('sheet')) || wbBonif.SheetNames[0];
    console.log('Using sheet:', sheetName);

    const wsBonif = wbBonif.Sheets[sheetName];
    const rawBonif = XLSX.utils.sheet_to_json(wsBonif, { defval: null, raw: true });
    console.log('Bonif rows:', rawBonif.length);

    const hBonif = rawBonif.length > 0 ? Object.keys(rawBonif[0]) : [];
    console.log('Bonif cols:', hBonif.slice(0, 10).join(', '));

    const bCodcli = normCol(hBonif, ['cod. parc', 'codparc', 'cod parc', 'cod_parc', 'codcli']);
    const bValor  = normCol(hBonif, ['fin.', 'fin', 'financeiro', 'vlr', 'valor']);
    const bData   = normCol(hBonif, ['emissao', 'emissão', 'emiss', 'data', 'dtped']);
    console.log('Bonif cols found:', { bCodcli, bValor, bData });

    // Bonificações: sem filtro de data mínima — incluir todo o histórico do arquivo
    const DATA_BONIF_MIN = new Date('2024-01-01');
    let totalBonif = 0;
    for (const r of rawBonif) {
      const dt = parseDate(r[bData]);
      if (!dt || dt < DATA_BONIF_MIN) continue;
      const valor = parseValor(r[bValor]);
      if (!valor || valor <= 0) continue;
      const anomes = dt.getFullYear() * 100 + (dt.getMonth() + 1);
      bonifCompact.push([String(r[bCodcli] || ''), valor, anomes]);
      totalBonif++;
    }
    console.log(`Bonificacoes loaded: ${totalBonif} registros`);
  } else {
    console.log('No bonificacao file found, skipping.');
  }
} catch(e) {
  console.warn('Bonificacao read error:', e.message);
}

// ---- OUTPUT ----
const output = {
  dados: dadosCompact,
  cadastro: cadCompact,
  produtos: prodCompact,
  metas: metasCompact,        // [codcli, metaValor, nomevend, mesRef]
  bonificacoes: bonifCompact, // [codcli, valor, anomes]
  meta: {
    colsDados: ['codcli','valor','anomes','descprod'],
    colsCadastro: ['codcli','nomecli','nomevend'],
    colsMetas: ['codcli','metaValor','nomevend','mesRef'],
    colsBonif: ['codcli','valor','anomes'],
    geradoEm: new Date().toISOString()
  }
};

const json = JSON.stringify(output);
fs.writeFileSync('/tmp/hanova/hanova_data.json', json);
console.log('JSON size:', (json.length / 1024 / 1024).toFixed(2), 'MB');

// Gzip + base64 para o dashboard
const zlib = require('zlib');
const compressed = zlib.gzipSync(Buffer.from(json, 'utf8'));
const b64 = compressed.toString('base64');
fs.writeFileSync('/tmp/hanova/hanova_data.b64', b64);
console.log('B64 size:', (b64.length / 1024).toFixed(0), 'KB');
console.log('Done!');
