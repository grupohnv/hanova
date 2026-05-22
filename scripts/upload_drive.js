const { execSync } = require('child_process');
const fs = require('fs');

const DADOS   = '/tmp/hanova/dados_relatorios';
const ENTRADA = '/tmp/hanova/entrada';

function rclone(args) {
  console.log('rclone', args);
  try { execSync('rclone ' + args, { stdio: 'inherit' }); }
  catch(e) { console.warn('Aviso:', e.message); }
}

// Upload dados processados atualizados
if (fs.existsSync(DADOS + '/dados_processados_powerbi_novo.xlsx')) {
  rclone('copy "' + DADOS + '/dados_processados_powerbi_novo.xlsx" "gdrive:dados relatorios/"');
  console.log('Dados processados enviados ao Drive');
}

// Upload bonif atualizado
const bonifFile = fs.readdirSync(DADOS).find(f =>
  f.toLowerCase().includes('onificad') && !f.startsWith('~') && (f.endsWith('.xlsx') || f.endsWith('.xls'))
);
if (bonifFile) {
  rclone('copy "' + DADOS + '/' + bonifFile + '" "gdrive:dados relatorios/"');
  console.log('Bonif enviado:', bonifFile);
}

// Upload do dashboard atualizado para o Google Drive
rclone('copy "./dashboard_hanova.html" "gdrive:aplicativos claude/"');
console.log('Dashboard enviado ao Drive: aplicativos claude/dashboard_hanova.html');

// Remover arquivos de entrada processados do Drive
const arquivosEntrada = fs.existsSync(ENTRADA)
  ? fs.readdirSync(ENTRADA).filter(f => (f.endsWith('.xlsx') || f.endsWith('.xls')) && !f.startsWith('~'))
  : [];

for (const f of arquivosEntrada) {
  if (!fs.existsSync(ENTRADA + '/' + f)) {
    rclone('deletefile "gdrive:dados relatorios/entrada/' + f + '"');
    console.log('Removido do Drive entrada:', f);
  }
}

console.log('Upload concluido!');
