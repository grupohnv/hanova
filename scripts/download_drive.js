const { execSync } = require('child_process');
const fs = require('fs');

const TMP     = '/tmp/hanova';
const DADOS   = TMP + '/dados_relatorios';
const METAS   = TMP + '/metas';
const ENTRADA = TMP + '/entrada';

[TMP, DADOS, METAS, ENTRADA].forEach(d => fs.mkdirSync(d, { recursive: true }));

function rclone(args) {
  console.log('rclone', args);
  execSync('rclone ' + args, { stdio: 'inherit' });
}

function rcloneSafe(args) {
  try {
    rclone(args);
  } catch(e) {
    console.warn('Aviso (nao fatal):', e.message.split('\n')[0]);
  }
}

// Arquivos de dados principais
rclone('copy "gdrive:dados relatorios/dados_processados_powerbi_novo.xlsx" "' + DADOS + '/"');
rclone('copy "gdrive:dados relatorios/cadastro-cliente-dimensao.xlsx" "' + DADOS + '/"');
rcloneSafe('copy "gdrive:dados relatorios/cadastro-cliente-dimensão.xlsx" "' + DADOS + '/"');
rclone('copy "gdrive:dados relatorios/Produtos Hanova.xlsx" "' + DADOS + '/"');

// Bonificacoes (qualquer arquivo com "bonificad" no nome)
rclone('copy "gdrive:dados relatorios/" "' + DADOS + '/" --include "*onificad*"');

// Pasta de entrada (novos relatorios a processar)
rcloneSafe('copy "gdrive:dados relatorios/entrada/" "' + ENTRADA + '/"');

// Metas — pasta pode ter nome com maiusculas variadas, tentar ambos
rcloneSafe('copy "gdrive:Comercial hanova/Metas/" "' + METAS + '/" --max-depth 2');
rcloneSafe('copy "gdrive:comercial hanova/metas/" "' + METAS + '/" --max-depth 2');
rcloneSafe('copy "gdrive:comercial hanova/Metas/" "' + METAS + '/" --max-depth 2');

// Listar arquivos baixados
console.log('\n=== Arquivos baixados ===');
execSync('find ' + TMP + ' -name "*.xlsx" -o -name "*.xls" | sort', { stdio: 'inherit' });
console.log('\nDownload concluido!');
