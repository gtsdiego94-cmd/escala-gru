const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

const REPO_DIR = 'C:\\Users\\User\\Desktop\\escala-gru';
const EXPORT_DIR = 'C:\\Export';
const DB_FILE = path.join(REPO_DIR, 'dados_vendas_mes.json');

// 1. Localizar arquivo de entrada (passado por argumento ou o mais recente em C:\Export)
let targetFile = process.argv[2];

if (!targetFile || !fs.existsSync(targetFile)) {
  if (fs.existsSync(EXPORT_DIR)) {
    const files = fs.readdirSync(EXPORT_DIR)
      .map(f => path.join(EXPORT_DIR, f))
      .filter(f => fs.statSync(f).isFile() && (f.endsWith('.csv') || f.endsWith('.xlsx') || f.endsWith('.txt')))
      .sort((a, b) => fs.statSync(b).mtime - fs.statSync(a).mtime);

    if (files.length > 0) {
      targetFile = files[0];
    }
  }
}

if (!targetFile || !fs.existsSync(targetFile)) {
  console.log('\n❌ Nenhum arquivo do OneBoss encontrado em C:\\Export.');
  console.log('👉 Exporte o relatório do OneBoss para C:\\Export ou arraste o arquivo para cima deste atalho.\n');
  process.exit(1);
}

console.log(`\n📂 Lendo relatório: ${path.basename(targetFile)}...`);

// 2. Parser CSV / XLSX
let textContent = '';
if (targetFile.endsWith('.xlsx') || targetFile.endsWith('.xls')) {
  // Se for Excel
  try {
    const XLSX = require(path.join(REPO_DIR, 'node_modules', 'xlsx'));
    const wb = XLSX.readFile(targetFile);
    textContent = XLSX.utils.sheet_to_csv(wb.Sheets[wb.SheetNames[0]]);
  } catch (e) {
    // Caso xlsx local não esteja instalado, tenta ler binário se for csv
    textContent = fs.readFileSync(targetFile, 'latin1');
  }
} else {
  textContent = fs.readFileSync(targetFile, 'latin1');
}

textContent = textContent.replace(/^\uFEFF/, '').trim();
const rawLines = textContent.split(/\r?\n/).map(l => l.trim()).filter(l => l.length > 0);

if (rawLines.length < 2) {
  console.log('❌ Arquivo vazio ou inválido.');
  process.exit(1);
}

// Detectar separador
let sep = ',';
const firstLine = rawLines[0];
if ((firstLine.match(/;/g) || []).length > (firstLine.match(/,/g) || []).length) sep = ';';
if ((firstLine.match(/\t/g) || []).length > (firstLine.match(/,/g) || []).length) sep = '\t';

function splitLine(line) {
  const regex = new RegExp('(?:^|' + (sep === ';' ? ';' : (sep === '\t' ? '\t' : ',')) + ')(?:"([^"]*(?:""[^"]*)*)"|([^"' + (sep === ';' ? ';' : (sep === '\t' ? '\t' : ',')) + ']*))', 'g');
  const matches = [];
  let match;
  while ((match = regex.exec(line)) !== null) {
    let val = match[1] !== undefined ? match[1].replace(/""/g, '"') : match[2];
    matches.push((val || '').trim());
  }
  return matches;
}

function norm(str) {
  return (str || '').normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().replace(/[^a-z0-9]/g, "");
}

// Localizar cabeçalho
let headerIdx = -1;
let idxValor = -1, idxFunc = -1, idxTipo = -1, idxTurno = -1, idxData = -1, idxRPS = -1;

for (let i = 0; i < Math.min(10, rawLines.length); i++) {
  const cols = splitLine(rawLines[i]).map(norm);
  const fIdx = cols.findIndex(h => h.includes('func') || h.includes('operador') || h.includes('vendedor') || h.includes('colab'));
  const vIdx = cols.findIndex(h => h === 'valor' || h === 'valortotal' || h === 'totalvenda' || h.includes('valor'));
  if (fIdx !== -1 && vIdx !== -1) {
    headerIdx = i;
    idxFunc = fIdx;
    idxValor = vIdx;
    idxTipo = cols.findIndex(h => h === 'tipo' || h === 'tiposervico' || h.includes('item'));
    idxTurno = cols.findIndex(h => h === 'turno' || h === 'turnoid');
    idxData = cols.findIndex(h => h.includes('dataturno') || h === 'data' || h.includes('emissao'));
    idxRPS = cols.findIndex(h => h.includes('rpsnumero') || h.includes('rps') || h.includes('numero'));
    break;
  }
}

if (headerIdx === -1) {
  console.log('❌ Não foi possível encontrar as colunas no relatório.');
  process.exit(1);
}

// 3. Carregar banco local de vendas acumuladas do mês
let bancoVendas = {};
if (fs.existsSync(DB_FILE)) {
  try {
    bancoVendas = JSON.parse(fs.readFileSync(DB_FILE, 'utf8'));
  } catch (e) {
    bancoVendas = {};
  }
}

let novasVendas = 0;
let dataUltima = '';

for (let i = headerIdx + 1; i < rawLines.length; i++) {
  const cols = splitLine(rawLines[i]);
  if (cols.length <= Math.max(idxValor, idxFunc)) continue;

  let rawVal = cols[idxValor] || '0';
  rawVal = rawVal.replace(/[^0-9,\.-]/g, '');
  if (rawVal.includes(',') && rawVal.includes('.')) rawVal = rawVal.replace(/\./g, '').replace(',', '.');
  else if (rawVal.includes(',')) rawVal = rawVal.replace(',', '.');
  const val = parseFloat(rawVal) || 0;

  const func = cols[idxFunc] ? cols[idxFunc].trim() : 'Não identificado';
  if (!func || norm(func) === 'total') continue;

  const tipo = norm(idxTipo !== -1 ? cols[idxTipo] : '');
  const isMala = tipo.includes('servico') || tipo.includes('serv') || tipo.includes('aerobag') || tipo.includes('lacre');
  const turno = idxTurno !== -1 && cols[idxTurno] ? cols[idxTurno].trim() : '-';
  const data = idxData !== -1 && cols[idxData] ? cols[idxData].trim() : new Date().toISOString().split('T')[0];
  const rps = idxRPS !== -1 && cols[idxRPS] ? cols[idxRPS].trim() : (i + '_' + Math.random());

  dataUltima = data;

  // Chave única para não duplicar venda se carregar o mesmo dia duas vezes
  const vendaKey = `${data}_${func}_${rps}_${val}_${tipo}`;
  if (!bancoVendas[vendaKey]) {
    bancoVendas[vendaKey] = { data, func, val, isMala, turno };
    novasVendas++;
  }
}

fs.writeFileSync(DB_FILE, JSON.stringify(bancoVendas, null, 2), 'utf8');
console.log(`✅ ${novasVendas} novos registros processados! Total no mês: ${Object.keys(bancoVendas).length} vendas.`);

// 4. Calcular estatísticas consolidadas do mês
let fatTotal = 0;
let malasTotal = 0;
const colabsMap = {};
const dailyMap = {};

Object.values(bancoVendas).forEach(v => {
  fatTotal += v.val;
  if (v.isMala) malasTotal++;

  // Por Colaborador
  const k = norm(v.func);
  if (!colabsMap[k]) {
    colabsMap[k] = { nome: v.func, turno: v.turno, fat: 0, malas: 0 };
  }
  colabsMap[k].fat += v.val;
  if (v.isMala) colabsMap[k].malas++;

  // Por Dia
  const diaNum = parseInt(v.data.split('-')[2] || v.data.split('/')[0]) || 1;
  dailyMap[diaNum] = (dailyMap[diaNum] || 0) + v.val;
});

const colabsList = Object.values(colabsMap)
  .sort((a, b) => b.fat - a.fat)
  .map((c, idx) => ({
    pos: idx + 1,
    nome: c.nome,
    turno: c.turno,
    fat: Math.round(c.fat),
    malas: c.malas,
    media: c.malas > 0 ? parseFloat((c.fat / c.malas).toFixed(2)) : 0
  }));

const dailyList = [];
for (let d = 1; d <= 30; d++) {
  if (dailyMap[d]) dailyList.push({ dia: d, val: Math.round(dailyMap[d]) });
}

function fmtBRL(v) {
  return 'R$ ' + (v || 0).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

console.log('\n📊 RESUMO DO MÊS ATUALIZADO:');
console.log(`💰 Faturamento Total: ${fmtBRL(fatTotal)}`);
console.log(`🧳 Total de Malas:    ${malasTotal}`);
console.log(`👥 Vendedores Ativos: ${colabsList.length}`);

// 5. Atualizar index.html e vendas.html
const indexPath = path.join(REPO_DIR, 'index.html');
let indexHtml = fs.readFileSync(indexPath, 'utf8');

// Injetar dados padrão em DADOS_ATUAIS se presente
const dadosInjetados = `
    let DADOS_ATUAIS = {
      faturamentoTotal: ${fatTotal},
      totalMalas: ${malasTotal},
      totalLinhas: ${Object.keys(bancoVendas).length},
      colabs: ${JSON.stringify(colabsList)},
      dataTurno: '${dataUltima || "Novembro/2026"}',
      arquivoNome: '${path.basename(targetFile)}'
    };
`;

indexHtml = indexHtml.replace(/let DADOS_ATUAIS = \{[\s\S]*?\};\n/, dadosInjetados);
fs.writeFileSync(indexPath, indexHtml, 'utf8');

// 6. Gerar texto para o WhatsApp
let zap = '✈️ *RESUMO DE VENDAS - ONEBOSS GRU*\n';
zap += `📅 *Atualizado até:* ${dataUltima || 'Hoje'}\n\n`;
zap += `💰 *Faturamento Total:* ${fmtBRL(fatTotal)}\n`;
zap += `🧳 *Total de Malas:* ${malasTotal}\n\n`;
zap += `🏆 *TOP 5 COLABORADORES:*\n`;
colabsList.slice(0, 5).forEach((c, i) => {
  const medal = i === 0 ? '🥇' : i === 1 ? '🥈' : i === 2 ? '🥉' : `${i + 1}º`;
  zap += `${medal} *${c.nome.toUpperCase()}* - ${fmtBRL(c.fat)} (${c.malas} malas)\n`;
});
zap += `\n🔗 *Acompanhe ao vivo:* https://gtsdiego94-cmd.github.io/escala-gru/\n`;
zap += `🚀 *Equipe Protecbag GRU*`;

// Copiar para clipboard no Windows
try {
  const zapFile = path.join(REPO_DIR, 'resumo_whatsapp.txt');
  fs.writeFileSync(zapFile, zap, 'utf8');
  execSync(`powershell.exe -Command "Get-Content -Path '${zapFile}' -Raw -Encoding UTF8 | Set-Clipboard"`);
  console.log('\n📲 Resumo copiado para a Área de Transferência! (Só dar Ctrl+V no WhatsApp)');
} catch (e) {}

// 7. Git commit automático se estiver configurado
try {
  console.log('\n🚀 Atualizando GitHub...');
  execSync(`git config user.name "gtsdiego94-cmd"`, { cwd: REPO_DIR });
  execSync(`git config user.email "diego@protecbag.com.br"`, { cwd: REPO_DIR });
  execSync(`git add .`, { cwd: REPO_DIR });
  execSync(`git commit -m "Atualização automática de vendas - ${dataUltima || 'Nov 2026'}"`, { cwd: REPO_DIR });
  console.log('✅ Commit realizado com sucesso!');
} catch (e) {
  console.log('ℹ️ Nenhuma alteração pendente para commit.');
}

console.log('\n🎉 PROCESSO CONCLUÍDO COM SUCESSO!\n');
