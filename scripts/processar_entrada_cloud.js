/**
 * HANOVA — Processador de Arquivos de Entrada
 *
 * Como usar:
 *   1. Salve o arquivo recebido em:
 *      G:/Meu Drive/dados relatorios/entrada/
 *      (qualquer nome, .xlsx ou .xls)
 *   2. Execute: node processar_entrada.js
 *
 * O script detecta automaticamente:
 *   - Se tem coluna PA → é Dados Processados
 *   - Se não tem coluna PA → é Bonificação
 *
 * Para cada arquivo:
 *   - Identifica o(s) mês(es) contido(s) nos dados
 *   - Remove esses meses do arquivo base
 *   - Insere os novos dados
 *   - Salva e apaga o arquivo de entrada
 */

const XLSX = require('xlsx');
const fs   = require('fs');
const path = require('path');

// ══════════════════════════════════════════════════
// CAMINHOS
// ══════════════════════════════════════════════════
const ENTRADA_DIR = '/tmp/hanova/entrada/';
const DADOS_PATH  = '/tmp/hanova/dados_relatorios/dados_processados_powerbi_novo.xlsx';
const BONIF_DIR   = '/tmp/hanova/dados_relatorios/';

// ══════════════════════════════════════════════════
// HELPERS
// ══════════════════════════════════════════════════
function log(msg) {
  const ts = new Date().toLocaleTimeString('pt-BR');
  console.log(`[${ts}] ${msg}`);
}

function normCol(headers, candidates) {
  return headers.find(h =>
    candidates.some(c => h.trim().toLowerCase() === c.toLowerCase() ||
                         h.trim().toLowerCase().includes(c.toLowerCase()))
  );
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

// "2026-05" — formato da coluna AnoMes no dados_processados
function toAnoMesStr(dt) {
  const y = dt.getFullYear();
  const m = String(dt.getMonth() + 1).padStart(2, '0');
  return `${y}-${m}`;
}

// Detecta os meses presentes nos dados a partir de uma coluna de data
function detectarMeses(rows, colData) {
  const meses = new Set();
  rows.forEach(r => {
    const dt = parseDate(r[colData]);
    if (dt) meses.add(toAnoMesStr(dt));
  });
  return meses;
}

// ══════════════════════════════════════════════════
// PROCESSAR DADOS PROCESSADOS
// ══════════════════════════════════════════════════
function processarDadosProcessados(inputPath) {
  log(`📊 Dados Processados — ${path.basename(inputPath)}`);

  // Ler arquivo de entrada
  const wbIn = XLSX.readFile(inputPath, { cellDates: true, raw: true });
  const wsIn = wbIn.Sheets[wbIn.SheetNames[0]];
  const rawIn = XLSX.utils.sheet_to_json(wsIn, { defval: null, raw: true });

  if (!rawIn.length) {
    log('  ⚠  Arquivo de entrada vazio, pulando.');
    return false;
  }

  const headersIn = Object.keys(rawIn[0]);
  const colPA      = headersIn.find(h => h.trim().toUpperCase() === 'PA');
  const colPeddata = normCol(headersIn, ['peddata','data','dtped','datpedido','dat_ped']);
  const colValorIn = normCol(headersIn, ['valor','vlr','vl_bruto','vl_liq','fin','financeiro']);
  const colNomeIn  = normCol(headersIn, ['nomecli','nome_cli','nomeparc','nome']);
  log(`  PA: ${colPA||'—'}  |  PedData: ${colPeddata}  |  Valor: ${colValorIn}  |  Nome: ${colNomeIn}`);

  if (!colPeddata) {
    log('  ❌ Coluna de data (peddata) não encontrada. Verifique o arquivo.');
    return false;
  }

  // 1. Filtrar linhas com PA vazio (somente se coluna PA existir)
  const antesPA = rawIn.length;
  const semPA = colPA
    ? rawIn.filter(r => {
        const v = r[colPA];
        return v !== null && v !== undefined && String(v).trim() !== '';
      })
    : rawIn;
  log(`  Linhas: ${antesPA} → após filtro PA: ${semPA.length} (${antesPA - semPA.length} removidas)`);

  // 2. Normalizar colunas e calcular AnoMes
  const COLUNAS_BASE = ['CODCLI','NOMECLI','DESCPROD','QTD','UFCLI','CIDADECLI','NOMEVEND','PedData','VALOR','AnoMes'];
  const processado = semPA.map(r => {
    const row = {};
    // Copiar apenas colunas conhecidas, normalizando nomes
    row['CODCLI']    = r['CODCLI'] ?? r['codcli'] ?? null;
    row['NOMECLI']   = colNomeIn ? r[colNomeIn] : null;
    row['DESCPROD']  = r['DESCPROD'] ?? r['descprod'] ?? null;
    row['QTD']       = r['QTD'] ?? r['qtd'] ?? null;
    row['UFCLI']     = r['UFCLI'] ?? r['ufcli'] ?? null;
    row['CIDADECLI'] = r['CIDADECLI'] ?? r['cidadecli'] ?? null;
    row['NOMEVEND']  = r['NOMEVEND'] ?? r['nomevend'] ?? null;
    row['PedData']   = r[colPeddata];
    row['VALOR']     = colValorIn ? r[colValorIn] : null;
    const dt = parseDate(r[colPeddata]);
    row['AnoMes']    = dt ? toAnoMesStr(dt) : null;
    return row;
  });

  // 3. Detectar meses presentes (AnoMes já é string "YYYY-MM", coleta direto)
  const mesesNovos = new Set(processado.map(r => r['AnoMes']).filter(Boolean));
  log(`  Meses nos novos dados: ${[...mesesNovos].sort().join(', ')}`);

  if (!mesesNovos.size) {
    log('  ❌ Não foi possível detectar meses. Verifique a coluna de data.');
    return false;
  }

  // 4. Abrir arquivo base
  const wbBase = XLSX.readFile(DADOS_PATH, { cellDates: true, raw: true });
  const wsBase = wbBase.Sheets[wbBase.SheetNames[0]];
  const rawBase = XLSX.utils.sheet_to_json(wsBase, { defval: null, raw: true });

  const headersBase  = rawBase.length > 0 ? Object.keys(rawBase[0]) : [];
  const colAnoMesBase = normCol(headersBase, ['anomes','ano_mes','ano mes','anoMes']);
  log(`  Coluna AnoMes na base: ${colAnoMesBase}`);

  // 5. Remover linhas dos meses sendo atualizados + normalizar coluna VALOR da base
  const colValorBase = normCol(headersBase, ['valor','vlr','vl_bruto','vl_liq']);
  const mantidos = (colAnoMesBase
    ? rawBase.filter(r => {
        const am = r[colAnoMesBase] ? String(r[colAnoMesBase]).trim() : '';
        return !mesesNovos.has(am);
      })
    : rawBase
  ).map(r => {
    // Normalizar: garantir coluna VALOR sem espaços
    const obj = { ...r };
    if (colValorBase && colValorBase !== 'VALOR') {
      obj['VALOR'] = r[colValorBase] ?? 0;
      delete obj[colValorBase];
    }
    return obj;
  });
  log(`  Base: ${rawBase.length} → após remover meses antigos: ${mantidos.length}`);

  // 6. Concatenar (remover linhas de total/resumo sem CODCLI) e salvar
  const final = [...mantidos, ...processado].filter(r => {
    const cod = r['CODCLI'];
    const desc = String(r['DESCPROD']||'').trim().toLowerCase();
    return cod && String(cod).trim() !== '' && desc !== 'total' && desc !== 'subtotal';
  });
  log(`  Total final: ${final.length} linhas`);

  const wsOut = XLSX.utils.json_to_sheet(final);
  wbBase.Sheets[wbBase.SheetNames[0]] = wsOut;
  XLSX.writeFile(wbBase, DADOS_PATH);
  log(`  ✅ Salvo: ${path.basename(DADOS_PATH)}`);
  return true;
}

// ══════════════════════════════════════════════════
// PROCESSAR BONIFICAÇÕES
// ══════════════════════════════════════════════════
function processarBonificacoes(inputPath) {
  log(`🎁 Bonificações — ${path.basename(inputPath)}`);

  // Encontrar arquivo de bonificações (suporte a múltiplos nomes)
  const allFiles = fs.readdirSync(BONIF_DIR);
  const bonifFile = allFiles.find(f =>
    f.toLowerCase().includes('bonificad') &&
    (f.endsWith('.xlsx') || f.endsWith('.xls')) &&
    !f.startsWith('~')
  );

  if (!bonifFile) {
    log('  ❌ Arquivo de bonificações não encontrado em ' + BONIF_DIR);
    return false;
  }

  const bonifPath = path.join(BONIF_DIR, bonifFile);
  log(`  Arquivo base: ${bonifFile}`);

  // Ler entrada
  const wbIn = XLSX.readFile(inputPath, { cellDates: true, raw: true });
  const wsIn = wbIn.Sheets[wbIn.SheetNames[0]];
  const rawIn = XLSX.utils.sheet_to_json(wsIn, { defval: null, raw: true });

  if (!rawIn.length) {
    log('  ⚠  Arquivo de entrada vazio, pulando.');
    return false;
  }

  const headersIn = Object.keys(rawIn[0]);
  const colDataIn = normCol(headersIn, ['emissao','emissão','emiss','data','dtped','emissao']);
  log(`  Coluna data entrada: ${colDataIn}`);

  if (!colDataIn) {
    log('  ❌ Coluna de data não encontrada. Verifique o arquivo.');
    return false;
  }

  // Meses nos novos dados
  const mesesNovos = detectarMeses(rawIn, colDataIn);
  log(`  Meses nos novos dados: ${[...mesesNovos].sort().join(', ')}`);

  // Ler base
  const wbBase = XLSX.readFile(bonifPath, { cellDates: true, raw: true });
  const sheetName = wbBase.SheetNames.find(s =>
    s.toLowerCase().includes('new') || s.toLowerCase().includes('sheet')
  ) || wbBase.SheetNames[0];

  const wsBase = wbBase.Sheets[sheetName];
  const rawBase = XLSX.utils.sheet_to_json(wsBase, { defval: null, raw: true });

  const headersBase = rawBase.length > 0 ? Object.keys(rawBase[0]) : [];
  const colDataBase = normCol(headersBase, ['emissao','emissão','emiss','data','dtped']);
  log(`  Aba base: "${sheetName}" | Coluna data: ${colDataBase} | Linhas: ${rawBase.length}`);

  // Detectar coluna de valor na base (pode ter espaços ex: " Fin. ")
  const colValorBase = normCol(headersBase, ['fin','financeiro','valor','vlr']);

  // Remover meses sendo atualizados + normalizar coluna Fin. da base
  const mantidos = (colDataBase
    ? rawBase.filter(r => {
        const dt = parseDate(r[colDataBase]);
        if (!dt) return true;
        return !mesesNovos.has(toAnoMesStr(dt));
      })
    : rawBase
  ).map(r => {
    const obj = { ...r };
    // Normalizar coluna de valor sem espaços
    if (colValorBase && colValorBase !== 'Fin.') {
      obj['Fin.'] = r[colValorBase] ?? 0;
      delete obj[colValorBase];
    }
    return obj;
  });
  log(`  Base: ${rawBase.length} → após remover meses antigos: ${mantidos.length}`);

  // Normalizar colunas do arquivo de entrada para coincidir com a base
  const colValorIn = normCol(headersIn, ['fin','financeiro','valor','vlr']);
  const normalizado = rawIn
    .filter(r => r[headersIn.find(h=>h.toLowerCase().includes('parc')||h.toLowerCase().includes('cod'))] && r[colDataIn])
    .map(r => {
      const obj = {};
      headersBase.forEach(hb => {
        // Tentar casar coluna da base com coluna da entrada (trim + lowercase)
        const match = headersIn.find(hi => hi.trim().toLowerCase() === hb.trim().toLowerCase());
        obj[hb] = match ? r[match] : null;
      });
      // Garantir valor normalizado
      if (colValorIn) obj['Fin.'] = r[colValorIn] ?? 0;
      return obj;
    });
  log(`  Entrada normalizada: ${normalizado.length} linhas`);

  // Concatenar e salvar
  const final = [...mantidos, ...normalizado];
  log(`  Total final: ${final.length} linhas`);

  const wsOut = XLSX.utils.json_to_sheet(final);
  wbBase.Sheets[sheetName] = wsOut;
  XLSX.writeFile(wbBase, bonifPath);
  log(`  ✅ Salvo: ${bonifFile}`);
  return true;
}

// ══════════════════════════════════════════════════
// MAIN
// ══════════════════════════════════════════════════
console.log('\n' + '═'.repeat(52));
console.log('  HANOVA — Processador de Entrada');
console.log('  ' + new Date().toLocaleString('pt-BR'));
console.log('═'.repeat(52));

if (!fs.existsSync(ENTRADA_DIR)) {
  console.error('❌ Pasta de entrada não encontrada:', ENTRADA_DIR);
  process.exit(1);
}

const arquivos = fs.readdirSync(ENTRADA_DIR).filter(f =>
  (f.endsWith('.xlsx') || f.endsWith('.xls')) && !f.startsWith('~')
);

if (!arquivos.length) {
  log('ℹ  Nenhum arquivo encontrado em "entrada/". Nada a fazer.');
  process.exit(0);
}

log(`📂 ${arquivos.length} arquivo(s) encontrado(s)`);
let processados = 0;

for (const arquivo of arquivos) {
  const fullPath = path.join(ENTRADA_DIR, arquivo);
  console.log('');

  try {
    // Detectar tipo: presença da coluna PA = dados processados
    const wb  = XLSX.readFile(fullPath, { cellDates: true, raw: true, sheetStubs: true });
    const ws  = wb.Sheets[wb.SheetNames[0]];
    const raw = XLSX.utils.sheet_to_json(ws, { defval: null, raw: true });
    const headers = raw.length > 0 ? Object.keys(raw[0]) : [];
    // Detectar tipo pelo nome da coluna de data:
    //   PedData  → dados processados (relatório de vendas)
    //   Emissão  → bonificações
    const temPedData = headers.some(h => h.trim().toLowerCase() === 'peddata');
    const temEmissao = headers.some(h => h.trim().toLowerCase().includes('emiss'));
    const temPA      = headers.some(h => h.trim().toUpperCase() === 'PA'); // fallback legado

    log(`  Detectado: PedData=${temPedData} | Emissão=${temEmissao} | PA=${temPA}`);
    let ok = false;
    if (temPedData && !temEmissao) {
      ok = processarDadosProcessados(fullPath);
    } else if (temEmissao) {
      ok = processarBonificacoes(fullPath);
    } else if (temPA) {
      // fallback legado: tem PA mas não tem PedData nem Emissão
      ok = processarDadosProcessados(fullPath);
    } else {
      log('  ❌ Não foi possível detectar o tipo do arquivo (sem PedData nem Emissão). Pulando.');
    }

    if (ok) {
      fs.unlinkSync(fullPath);
      log(`  🗑️  Arquivo removido da pasta de entrada`);
      processados++;
    }

  } catch (err) {
    log(`  ❌ Erro em "${arquivo}": ${err.message}`);
    console.error(err);
  }
}

console.log('');
console.log('═'.repeat(52));
if (processados > 0) {
  log(`✅ ${processados} arquivo(s) processado(s) com sucesso!`);
} else {
  log('⚠  Nenhum arquivo foi processado (verifique os erros acima).');
  process.exit(1);
}
console.log('═'.repeat(52) + '\n');
