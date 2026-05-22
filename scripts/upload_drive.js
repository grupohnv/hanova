const { execSync } = require('child_process');
const fs = require('fs');

const DADOS   = '/tmp/hanova/dados_relatorios';
const ENTRADA = '/tmp/hanova/entrada';

function rclone(args) {
  console.log('rclone', args);
  try { execSync('rclone ' + args, { stdio: 'inherit' }); }
  catch(e) { console.warn('Aviso:', e.message.split('\n')[0]); }
}

// ── Upload dados processados atualizados ──────────────────
if (fs.existsSync(DADOS + '/dados_processados_powerbi_novo.xlsx')) {
  rclone('copy "' + DADOS + '/dados_processados_powerbi_novo.xlsx" "gdrive:dados relatorios/"');
  console.log('Dados processados enviados ao Drive');
}

// ── Upload bonif atualizado ───────────────────────────────
const bonifFile = fs.readdirSync(DADOS).find(f =>
  f.toLowerCase().includes('onificad') && !f.startsWith('~') && (f.endsWith('.xlsx') || f.endsWith('.xls'))
);
if (bonifFile) {
  rclone('copy "' + DADOS + '/' + bonifFile + '" "gdrive:dados relatorios/"');
  console.log('Bonif enviado:', bonifFile);
}

// ── Upload do dashboard atualizado para o Google Drive ────
rclone('copy "./dashboard_hanova.html" "gdrive:aplicativos claude/"');
console.log('Dashboard enviado ao Drive: aplicativos claude/dashboard_hanova.html');

// ── Limpar arquivos de entrada do Drive ───────────────────
// Estratégia: listar arquivos na pasta entrada do Drive,
// e deletar os que não estão mais na pasta local (foram processados com sucesso)
try {
  console.log('\nVerificando arquivos de entrada no Drive...');

  // Listar arquivos no Drive entrada via rclone lsf
  let driveFiles = [];
  try {
    const out = execSync('rclone lsf "gdrive:dados relatorios/entrada/" --files-only', { encoding: 'utf8' });
    driveFiles = out.split('\n').map(f => f.trim()).filter(f => f && (f.endsWith('.xlsx') || f.endsWith('.xls')));
    console.log('Arquivos no Drive entrada:', driveFiles.length, driveFiles);
  } catch(e) {
    console.log('Pasta entrada do Drive vazia ou inacessível.');
    driveFiles = [];
  }

  // Arquivos ainda na pasta local (não processados / falharam)
  const localFiles = fs.existsSync(ENTRADA)
    ? fs.readdirSync(ENTRADA).filter(f => (f.endsWith('.xlsx') || f.endsWith('.xls')) && !f.startsWith('~'))
    : [];
  console.log('Arquivos locais restantes:', localFiles);

  // Deletar do Drive todos que foram processados (não estão mais no local)
  for (const f of driveFiles) {
    if (!localFiles.includes(f)) {
      console.log('Deletando do Drive (processado):', f);
      rclone('deletefile "gdrive:dados relatorios/entrada/' + f.replace(/"/g, '\"') + '"');
    } else {
      console.log('Mantendo no Drive (não processado):', f);
    }
  }
} catch(e) {
  console.warn('Aviso na limpeza de entrada:', e.message);
}

console.log('\nUpload concluido!');
