const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');
const XLSX = require('xlsx');

const REPO_DIR = 'C:\\Users\\User\\Desktop\\escala-gru';
const EXPORT_DIR = 'C:\\Export';
const DB_FILE = path.join(REPO_DIR, 'dados_vendas_mes.json');

// Mapeamento de Turno ID do OneBoss para Turno A, B, D
// Turno 1 (001) = Turno A (Manhã)
// Turno 2 (002) = Turno B (Tarde)
// Turno 3 (003) = Turno D (Madrugada - antes da meia noite)
// Turno 4 (004) = Turno D (Madrugada - depois da meia noite)
function turnoIdParaLetra(turnoId) {
  const t = (turnoId || '').trim();
  if (t === '001' || t === '1') return 'A';
  if (t === '002' || t === '2') return 'B';
  if (t === '003' || t === '3') return 'D';
  if (t === '004' || t === '4') return 'D';
  return null; // não identificado pelo ID
}

function inferirTurno(nome, turnoIdStr, turnoNumStr) {
  const normNome = (nome || '').toUpperCase().trim();

  // 1. Prioridade: Turno ID do OneBoss (001, 002, 003, 004)
  const porId = turnoIdParaLetra(turnoIdStr);
  if (porId) return porId;

  // 2. Se nome tem d2 = Turno D (Madrugada)
  if (normNome.endsWith('D2') || normNome.includes(' D2') || normNome.includes('(D2)')) {
    return 'D';
  }

  // 3. Número do turno simples
  if (turnoNumStr === '1') return 'A';
  if (turnoNumStr === '2') return 'B';
  if (turnoNumStr === '3' || turnoNumStr === '4') return 'D';

  // 4. Fallback: B
  return 'B';
}

// 1. Localizar arquivo de entrada
let targetFile = process.argv[2];
if (!targetFile || !fs.existsSync(targetFile)) {
  if (fs.existsSync(EXPORT_DIR)) {
    const files = fs.readdirSync(EXPORT_DIR)
      .map(f => path.join(EXPORT_DIR, f))
      .filter(f => fs.statSync(f).isFile() && (f.endsWith('.csv') || f.endsWith('.xlsx') || f.endsWith('.txt')))
      .sort((a, b) => fs.statSync(b).mtime - fs.statSync(a).mtime);

    if (files.length > 0) targetFile = files[0];
  }
}

if (!targetFile || !fs.existsSync(targetFile)) {
  console.log('\n❌ Nenhum arquivo do OneBoss encontrado em C:\\Export.');
  console.log('👉 Salve o relatório do OneBoss em C:\\Export e clique novamente.\n');
  process.exit(1);
}

console.log(`\n📂 Lendo relatório: ${path.basename(targetFile)}...`);

// Lê o arquivo — suporta XLSX e CSV/TXT
let rawLines = [];

const ext = path.extname(targetFile).toLowerCase();
if (ext === '.xlsx' || ext === '.xls') {
  // Lê XLSX com SheetJS (node)
  const wb = XLSX.readFile(targetFile);
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  rawLines = rows.map(row => row.map(c => String(c).trim()).join('\t')).filter(l => l.trim().length > 0);
  // Força separador tab para linhas geradas do XLSX
  var forceSep = '\t';
} else {
  // CSV/TXT
  let textContent = fs.readFileSync(targetFile, 'latin1').replace(/^\uFEFF/, '').trim();
  rawLines = textContent.split(/\r?\n/).map(l => l.trim()).filter(l => l.length > 0);
  var forceSep = null;
}

if (rawLines.length < 2) {
  console.log('❌ Arquivo vazio ou sem linhas de dados.');
  process.exit(1);
}

let sep = forceSep || ',';
if (!forceSep) {
  const firstLine = rawLines[0];
  if ((firstLine.match(/;/g) || []).length > (firstLine.match(/,/g) || []).length) sep = ';';
  if ((firstLine.match(/\t/g) || []).length > (firstLine.match(/,/g) || []).length) sep = '\t';
}

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

let headerIdx = -1;
let idxValor = -1, idxFunc = -1, idxTipo = -1, idxTurnoNum = -1, idxTurnoId = -1, idxData = -1, idxHora = -1, idxRPS = -1;

for (let i = 0; i < Math.min(10, rawLines.length); i++) {
  const cols = splitLine(rawLines[i]).map(norm);
  const fIdx = cols.findIndex(h => h.includes('func') || h.includes('operador') || h.includes('vendedor') || h.includes('colab'));
  const vIdx = cols.findIndex(h => h === 'valor' || h === 'valortotal' || h === 'totalvenda' || h.includes('valor'));
  if (fIdx !== -1 && vIdx !== -1) {
    headerIdx = i;
    idxFunc = fIdx;
    idxValor = vIdx;
    idxTipo = cols.findIndex(h => h === 'tipo' || h === 'tiposervico' || h.includes('item'));
    // Separa "Turno" (número 1,2,3,4) de "Turno ID" (001,002,003,004)
    idxTurnoId  = cols.findIndex(h => h === 'turnoid');
    idxTurnoNum = cols.findIndex(h => h === 'turno' && !cols[cols.indexOf(h)].includes('id'));
    if (idxTurnoNum === -1) idxTurnoNum = cols.findIndex(h => h === 'turno');
    idxData = cols.findIndex(h => h.includes('dataturno') || h === 'data' || h.includes('emissao'));
    idxHora = cols.findIndex(h => h === 'hora' || h.includes('horario'));
    idxRPS = cols.findIndex(h => h.includes('rpsnumero') || h.includes('rps') || h.includes('numero'));
    console.log(`📋 Colunas detectadas: Func=${idxFunc} Valor=${idxValor} TurnoNum=${idxTurnoNum} TurnoID=${idxTurnoId} Tipo=${idxTipo} Data=${idxData}`);
    break;
  }
}

if (headerIdx === -1) {
  console.log('❌ Colunas do OneBoss não identificadas.');
  process.exit(1);
}

// Banco acumulado
let bancoVendas = {};
if (fs.existsSync(DB_FILE)) {
  try { bancoVendas = JSON.parse(fs.readFileSync(DB_FILE, 'utf8')); } catch (e) { bancoVendas = {}; }
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
  const rawTurnoId  = idxTurnoId  !== -1 && cols[idxTurnoId]  ? cols[idxTurnoId].trim()  : '';
  const rawTurnoNum = idxTurnoNum !== -1 && cols[idxTurnoNum] ? cols[idxTurnoNum].trim() : '';
  const hora = idxHora !== -1 && cols[idxHora] ? cols[idxHora].trim() : '';
  const data = idxData !== -1 && cols[idxData] ? cols[idxData].trim() : new Date().toISOString().split('T')[0];
  const rps = idxRPS !== -1 && cols[idxRPS] ? cols[idxRPS].trim() : (i + '_' + Math.random());

  const turnoDefinido = inferirTurno(func, rawTurnoId, rawTurnoNum);
  dataUltima = data;

  const vendaKey = `${data}_${func}_${rps}_${val}_${tipo}`;
  if (!bancoVendas[vendaKey]) {
    bancoVendas[vendaKey] = { data, func, val, isMala, turno: turnoDefinido };
    novasVendas++;
  }
}

fs.writeFileSync(DB_FILE, JSON.stringify(bancoVendas, null, 2), 'utf8');
console.log(`✅ ${novasVendas} novos registros! Total acumulado no mês: ${Object.keys(bancoVendas).length} vendas.`);

// Totais e Turnos
let fatTotal = 0;
let malasTotal = 0;
const turnosTotais = {
  A: { fat: 0, malas: 0 },
  B: { fat: 0, malas: 0 },
  D: { fat: 0, malas: 0 }
};

const colabsMap = {};
const dailyMap = {};

Object.values(bancoVendas).forEach(v => {
  fatTotal += v.val;
  if (v.isMala) malasTotal++;

  const t = (v.turno === 'A' || v.turno === 'B' || v.turno === 'D') ? v.turno : 'B';
  turnosTotais[t].fat += v.val;
  if (v.isMala) turnosTotais[t].malas++;

  const k = norm(v.func);
  if (!colabsMap[k]) {
    colabsMap[k] = { nome: v.func, turno: t, fat: 0, malas: 0 };
  }
  colabsMap[k].fat += v.val;
  if (v.isMala) colabsMap[k].malas++;

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

function fmtBRL(v) {
  return 'R$ ' + (v || 0).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

console.log('\n📊 DESEMPENHO OFICIAL POR TURNO:');
console.log(`🌅 Turno A (Manhã):     ${fmtBRL(turnosTotais.A.fat)} (${turnosTotais.A.malas} malas)`);
console.log(`☀️ Turno B (Tarde):     ${fmtBRL(turnosTotais.B.fat)} (${turnosTotais.B.malas} malas)`);
console.log(`🌙 Turno D (Madrugada): ${fmtBRL(turnosTotais.D.fat)} (${turnosTotais.D.malas} malas)`);
console.log(`💰 Faturamento Geral:   ${fmtBRL(fatTotal)}`);

// Atualizar index.html com dados por turno
const indexPath = path.join(REPO_DIR, 'index.html');
let indexHtml = fs.readFileSync(indexPath, 'utf8');

const dadosInjetados = `
    let DADOS_ATUAIS = {
      faturamentoTotal: ${fatTotal},
      totalMalas: ${malasTotal},
      totalLinhas: ${Object.keys(bancoVendas).length},
      turnos: ${JSON.stringify(turnosTotais)},
      colabs: ${JSON.stringify(colabsList)},
      dataTurno: '${dataUltima || "Novembro/2026"}',
      arquivoNome: '${path.basename(targetFile)}'
    };
`;

indexHtml = indexHtml.replace(/let DADOS_ATUAIS = \{[\s\S]*?\};\n/, dadosInjetados);
fs.writeFileSync(indexPath, indexHtml, 'utf8');

// Montar resumo formatado do WhatsApp separado por turno
let zap = '✈️ *RESUMO DE VENDAS - ONEBOSS GRU*\n';
zap += `📅 *Atualizado:* ${dataUltima || 'Hoje'}\n\n`;
zap += `💰 *FATURAMENTO TOTAL:* ${fmtBRL(fatTotal)} (${malasTotal} malas)\n\n`;

zap += `⏱️ *DESEMPENHO POR TURNO:*\n`;
const pctA = fatTotal > 0 ? ((turnosTotais.A.fat / fatTotal) * 100).toFixed(1) : '0';
const pctB = fatTotal > 0 ? ((turnosTotais.B.fat / fatTotal) * 100).toFixed(1) : '0';
const pctD = fatTotal > 0 ? ((turnosTotais.D.fat / fatTotal) * 100).toFixed(1) : '0';

zap += `🌅 *TURNO A (Manhã):* ${fmtBRL(turnosTotais.A.fat)} (${turnosTotais.A.malas} malas) • ${pctA}%\n`;
zap += `☀️ *TURNO B (Tarde):* ${fmtBRL(turnosTotais.B.fat)} (${turnosTotais.B.malas} malas) • ${pctB}%\n`;
zap += `🌙 *TURNO D (Madrugada):* ${fmtBRL(turnosTotais.D.fat)} (${turnosTotais.D.malas} malas) • ${pctD}%\n\n`;

zap += `🏆 *DESTAQUES POR TURNO:*\n`;
const topA = colabsList.filter(c => c.turno === 'A').slice(0, 2);
const topB = colabsList.filter(c => c.turno === 'B').slice(0, 2);
const topD = colabsList.filter(c => c.turno === 'D').slice(0, 2);

if (topA.length) zap += `*Turno A:* 🥇 ${topA[0].nome} (${fmtBRL(topA[0].fat)})\n`;
if (topB.length) zap += `*Turno B:* 🥇 ${topB[0].nome} (${fmtBRL(topB[0].fat)})\n`;
if (topD.length) zap += `*Turno D:* 🥇 ${topD[0].nome} (${fmtBRL(topD[0].fat)})\n`;

zap += `\n🔗 *Acompanhe ao vivo:* https://gtsdiego94-cmd.github.io/escala-gru/\n`;
zap += `🚀 *Equipe Protecbag GRU*`;

try {
  const zapFile = path.join(REPO_DIR, 'resumo_whatsapp.txt');
  fs.writeFileSync(zapFile, zap, 'utf8');
  execSync(`powershell.exe -Command "Get-Content -Path '${zapFile}' -Raw -Encoding UTF8 | Set-Clipboard"`);
  console.log('\n📲 Resumo com divisão por Turnos copiado para a Área de Transferência!');
} catch (e) {}

// Git Commit e Push
try {
  console.log('\n🚀 Atualizando GitHub...');
  execSync(`git config user.name "gtsdiego94-cmd"`, { cwd: REPO_DIR });
  execSync(`git config user.email "diego@protecbag.com.br"`, { cwd: REPO_DIR });
  execSync(`git add .`, { cwd: REPO_DIR });
  execSync(`git commit -m "Atualização de vendas com divisão por turno"`, { cwd: REPO_DIR });
  console.log('✅ Commit realizado localmente.');
  
  // Tenta o push automático
  try {
    execSync(`git push origin main`, { cwd: REPO_DIR, stdio: 'ignore' });
    console.log('🌐 GitHub Pages sincronizado com sucesso!');
  } catch (pushErr) {
    console.log('ℹ️ Para sincronizar online pela primeira vez, rode: git push origin main');
  }
} catch (e) {}

console.log('\n🎉 PROCESSO CONCLUÍDO COM SUCESSO!\n');
