const { execSync } = require('child_process');
const fs = require('fs');

const TMP     = '/tmp/hanova';
const DADOS   = TMP + '/dados_relatorios';
const METAS   = TMP + '/metas';
const ENTRADA = TMP + '/entrada';

[TMP, DADOS, METAS, ENTRADA].forEach(d => fs.mkdirSync(d, { recursive: true }));

function rclone(cmd) {
  console.log('rclone', cmd);
  execSync('rclone ' + cmd, { stdio: 'inherit' });
}

function rcloneSafe(cmd) {
  try { rclone(cmd); } catch(e) { console.warn('Aviso (nao fatal):', e.message.split('\n')[0]); }
}

// Arquivos de dados principais
rclone('copy "gdrive:dados relatorios/dados_processados_powerbi_novo.xlsx" "' + DADOS + '/"');
rclone('copy "gdrive:dados relatorios/Produtos Hanova.xlsx" "' + DADOS + '/"');

// Cadastro cliente — usa --include para lidar com caracteres especiais no nome
rclone('copy "gdrive:dados relatorios/" "' + DADOS + '/" --include "cadastro-cliente*"');

// Bonificacoes
rclone('copy "gdrive:dados relatorios/" "' + DADOS + '/" --include "*onificad*"');

// Pasta de entrada
rcloneSafe('copy "gdrive:dados relatorios/entrada/" "' + ENTRADA + '/"');

// Metas — tentar variações de maiusculas/minusculas
rcloneSafe('copy "gdrive:Comercial hanova/Metas/" "' + METAS + '/" --max-depth 2');
rcloneSafe('copy "gdrive:comercial hanova/metas/" "' + METAS + '/" --max-depth 2');

// Listar arquivos baixados
console.log('\n=== Arquivos baixados ===');
execSync('find ' + TMP + ' -name "*.xlsx" -o -name "*.xls" | sort', { stdio: 'inherit' });
console.log('\nDownload concluido!');
