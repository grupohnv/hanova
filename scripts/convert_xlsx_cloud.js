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
// Tentar ambos os nomes (com e sem acento)
let wbCad;
const cadFiles = fs.readdirSync('/tmp/hanova/dados_relatorios/')
  .filter(f => f.toLowerCase().startsWith('cadastro-cliente') && (f.endsWith('.xlsx') || f.endsWith('.xls')));
console.log('Cad files found:', cadFiles);
wbCad = XLSX.readFile('/tmp/hanova/dados_relatorios/' + (cadFiles[0] || 'cadastro-cliente-dimensao.xlsx'), { cellDates: true, raw: true });
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
// Lê todos os arquivos .xlsx recursivamente em /tmp/hanova/metas/
// Suporta tanto a estrutura antiga (subpastas MM-YYYY com colMap na linha 0)
// quanto planilhas normais com header direto
const metasCompact = [];  // [codcli, metaValor]

function readAllXlsxFiles(dir) {
  const result = [];
  try {
    const entries = fs.readdirSync(dir, { withFileTypes: true });
    for (const e of entries) {
      const full = path.join(dir, e.name);
      if (e.isDirectory()) {
        result.push(...readAllXlsxFiles(full));
      } else if ((e.name.endsWith('.xlsx') || e.name.endsWith('.xls')) &&
                 !e.name.startsWith('~') && !e.name.startsWith('.')) {
        result.push(full);
      }
    }
  } catch(e) { /* ignore */ }
  return result;
}

function tryParseMeta(filePath) {
  const wb = XLSX.readFile(filePath, { raw: true });
  const results = [];

  for (const sheetName of wb.SheetNames) {
    const ws = wb.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(ws, { defval: null, raw: true });
    if (rows.length < 2) continue;

    const firstRowKeys = Object.keys(rows[0]);

    // Método 1: cabeçalho direto — procura colunas por nome nas chaves do objeto
    let kCodcli = normCol(firstRowKeys, ['codcli','cod_cli','codparc','codigo','cod.cli','cod cli']);
    let kMeta   = normCol(firstRowKeys, ['validar meta','meta','vlr meta','valor meta','meta_valor','metavalor','vl_meta']);
    let kNomevend = normCol(firstRowKeys, ['nomevend','nome_vend','vendedor','representante','rep']);

    if (kCodcli && kMeta) {
      console.log('  Meta (direto):', path.basename(filePath), 'sheet:', sheetName, '->', kCodcli, kMeta);
      for (const r of rows) {
        const codcli = String(r[kCodcli] || '').trim();
        const metaVal = typeof r[kMeta] === 'number' ? r[kMeta] : parseValor(r[kMeta]);
        if (codcli && metaVal > 0) {
          results.push([codcli, metaVal]);
        }
      }
      break; // Achou nesta aba, não precisa tentar as outras
    }

    // Método 2: linha 0 contém mapeamento de colunas (formato especial)
    // Os valores da linha 0 são os nomes reais das colunas
    const colMap = rows[0];
    const colMapVals = Object.values(colMap).map(v => String(v || '').trim().toUpperCase());
    const hasCodcli = colMapVals.some(v => v.includes('CODCLI') || v.includes('COD_CLI') || v.includes('CODPARC'));
    const hasMeta   = colMapVals.some(v => v.includes('META') || v.includes('VALIDAR'));

    if (hasCodcli && hasMeta) {
      const mapKeys = Object.keys(colMap);
      kCodcli   = mapKeys.find(k => {const v=String(colMap[k]||'').toUpperCase(); return v.includes('CODCLI')||v.includes('COD_CLI')||v.includes('CODPARC');});
      kMeta     = mapKeys.find(k => {const v=String(colMap[k]||'').toUpperCase(); return v.includes('VALIDAR META')||v==='META';});
      if (!kMeta) kMeta = mapKeys.find(k => {const v=String(colMap[k]||'').toUpperCase(); return v.includes('META');});
      kNomevend = mapKeys.find(k => {const v=String(colMap[k]||'').toUpperCase(); return v.includes('NOMEVEND')||v.includes('VENDEDOR')||v.includes('REPRESENTANTE');});

      if (kCodcli && kMeta) {
        console.log('  Meta (colmap):', path.basename(filePath), 'sheet:', sheetName, '->', kCodcli, kMeta);
        for (const r of rows.slice(1)) {
          const codcli = String(r[kCodcli] || '').trim();
          const metaVal = typeof r[kMeta] === 'number' ? r[kMeta] : parseValor(r[kMeta]);
          if (codcli && metaVal > 0) {
            results.push([codcli, metaVal]);
          }
        }
        break;
      }
    }
  }
  return results;
}

try {
  const META_BASE = '/tmp/hanova/metas/';
  const metaFiles = readAllXlsxFiles(META_BASE);
  console.log('Meta files found:', metaFiles.length, metaFiles.map(f => path.relative(META_BASE, f)));

  for (const filePath of metaFiles) {
    try {
      const rows = tryParseMeta(filePath);
      metasCompact.push(...rows);
      if (rows.length > 0) console.log('  Loaded', rows.length, 'metas from', path.basename(filePath));
    } catch(e) {
      console.warn('  Skip', path.basename(filePath), e.message);
    }
  }
  console.log('Total metas:', metasCompact.length);
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
    const bonifFile = bonifFiles.sort().reverse()[0];
    console.log('Reading bonificacoes:', bonifFile);
    const wbBonif = XLSX.readFile(path.join(dadosDir, bonifFile), { cellDates: true, raw: true });

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
    console.log('Bonificacoes loaded:', totalBonif, 'registros');
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
  metas: metasCompact,        // [codcli, metaValor]
  bonificacoes: bonifCompact, // [codcli, valor, anomes]
  meta: {
    colsDados: ['codcli','valor','anomes','descprod'],
    colsCadastro: ['codcli','nomecli','nomevend'],
    colsMetas: ['codcli','metaValor'],
    colsBonif: ['codcli','valor','anomes'],
    geradoEm: new Date().toISOString()
  }
};

const json = JSON.stringify(output);
fs.writeFileSync('/tmp/hanova/hanova_data.json', json);
console.log('JSON size:', (json.length / 1024 / 1024).toFixed(2), 'MB');

const zlib = require('zlib');
const compressed = zlib.gzipSync(Buffer.from(json, 'utf8'));
const b64 = compressed.toString('base64');
fs.writeFileSync('/tmp/hanova/hanova_data.b64', b64);
console.log('B64 size:', (b64.length / 1024).toFixed(0), 'KB');
console.log('Done!');
