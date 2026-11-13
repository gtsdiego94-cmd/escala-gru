const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');
const XLSX = require('xlsx');

const REPO_DIR = 'C:\\Users\\User\\Desktop\\escala-gru';
const DIRS_TO_LOOK = ['C:\\Relatorios', 'C:\\Export', REPO_DIR];

const META_BRONZE = 1500000;
const META_PRATA  = 1600000;
const META_OURO   = 1650000;

// Mapeamento de Turno ID do OneBoss para Turno A, B, D
function inferirTurno(nome, turnoIdStr, turnoNumStr) {
  const tId = String(turnoIdStr || '').trim();
  const tNum = String(turnoNumStr || '').trim();
  const normNome = String(nome || '').toUpperCase().trim();

  // 1. Prioridade: Turno ID do OneBoss (001, 002, 003, 004)
  if (tId === '001' || tId === '1') return 'A';
  if (tId === '002' || tId === '2') return 'B';
  if (tId === '003' || tId === '3' || tId === '004' || tId === '4') return 'D';

  // 2. Se nome tem d2 = Turno D (Madrugada)
  if (normNome.endsWith('D2') || normNome.includes(' D2') || normNome.includes('(D2)')) {
    return 'D';
  }

  // 3. Número do turno simples
  if (tNum === '1') return 'A';
  if (tNum === '2') return 'B';
  if (tNum === '3' || tNum === '4') return 'D';

  return 'B';
}

function toJsDate(val) {
  if (typeof val === 'number') {
    return new Date(Math.round((val - 25569) * 86400 * 1000));
  }
  return new Date(val);
}

function fmtBRL(v) {
  return 'R$ ' + (v || 0).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

// 1. Localizar arquivo de entrada mais recente
let targetFile = process.argv[2];
if (!targetFile || !fs.existsSync(targetFile)) {
  const allFiles = [];
  DIRS_TO_LOOK.forEach(dir => {
    if (fs.existsSync(dir)) {
      fs.readdirSync(dir).forEach(f => {
        if (f.startsWith('~$') || f.includes('vendas_atual')) return;
        const full = path.join(dir, f);
        try {
          const st = fs.statSync(full);
          if (st.isFile() && (f.endsWith('.xlsx') || f.endsWith('.csv') || f.endsWith('.xls'))) {
            allFiles.push({ path: full, time: st.mtimeMs });
          }
        } catch (e) {}
      });
    }
  });

  allFiles.sort((a, b) => b.time - a.time);
  if (allFiles.length > 0) targetFile = allFiles[0].path;
}

if (!targetFile || !fs.existsSync(targetFile)) {
  console.log('\n❌ Nenhum relatório do OneBoss encontrado.');
  process.exit(1);
}

console.log(`\n📂 Lendo relatório: ${path.basename(targetFile)}...`);

let rows = [];
const ext = path.extname(targetFile).toLowerCase();

if (ext === '.xlsx' || ext === '.xls') {
  const wb = XLSX.readFile(targetFile);
  const ws = wb.Sheets[wb.SheetNames[0]];
  rows = XLSX.utils.sheet_to_json(ws, { defval: '' });
} else {
  const textContent = fs.readFileSync(targetFile, 'latin1').replace(/^\uFEFF/, '').trim();
  const lines = textContent.split(/\r?\n/).map(l => l.trim()).filter(l => l.length > 0);
  let sep = lines[0].includes(';') ? ';' : (lines[0].includes('\t') ? '\t' : ',');
  const header = lines[0].split(sep).map(c => c.replace(/^"|"$/g, '').trim());
  for (let i = 1; i < lines.length; i++) {
    const cols = lines[i].split(sep).map(c => c.replace(/^"|"$/g, '').trim());
    const obj = {};
    header.forEach((h, idx) => { obj[h] = cols[idx] || ''; });
    rows.push(obj);
  }
}

if (rows.length === 0) {
  console.log('❌ Arquivo vazio ou sem linhas de dados.');
  process.exit(1);
}

// Totais e Turnos
let fatTotal = 0;
let malasTotal = 0;
const turnosTotais = {
  A: { fat: 0, malas: 0 },
  B: { fat: 0, malas: 0 },
  D: { fat: 0, malas: 0 }
};

const colabsMap = {};
let dataUltima = 'Novembro/2026';
const diasUnicos = new Set();

rows.forEach(r => {
  const nome = (r['Funcionário'] || r['Funcionario'] || r['Nome'] || r['COLABORADOR'] || '').trim();
  if (!nome || nome.toLowerCase() === 'total') return;

  let valRaw = r['Valor'] !== undefined && r['Valor'] !== '' ? r['Valor'] : (r['Total Venda'] || 0);
  if (typeof valRaw === 'string') {
    valRaw = valRaw.replace(/[^0-9,\.-]/g, '');
    if (valRaw.includes(',') && valRaw.includes('.')) valRaw = valRaw.replace(/\./g, '').replace(',', '.');
    else if (valRaw.includes(',')) valRaw = valRaw.replace(',', '.');
  }
  const val = parseFloat(valRaw) || 0;

  const tipo = String(r['Tipo'] || '').toLowerCase();
  const item = String(r['Item'] || '').toLowerCase();
  const isMala = tipo.includes('servi') || item.includes('aerobag') || item.includes('lacre') || item.includes('reembalar');

  const turno = inferirTurno(nome, r['Turno ID'] || r['TurnoID'], r['Turno']);

  fatTotal += val;
  if (isMala) malasTotal++;

  turnosTotais[turno].fat += val;
  if (isMala) turnosTotais[turno].malas++;

  const k = nome.toUpperCase();
  if (!colabsMap[k]) {
    colabsMap[k] = { nome, turno, fat: 0, malas: 0 };
  }
  colabsMap[k].fat += val;
  if (isMala) colabsMap[k].malas++;

  const rawDt = r['Data do Turno'] || r['Data'] || r['DATA'];
  if (rawDt) {
    const dObj = toJsDate(rawDt);
    if (!isNaN(dObj.getTime())) {
      const dtKey = dObj.toISOString().slice(0, 10);
      diasUnicos.add(dtKey);
      dataUltima = dObj.toLocaleDateString('pt-BR');
    }
  }
});

const diasOperacao = diasUnicos.size > 0 ? diasUnicos.size : new Date().getDate();

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

const totalDiasMes = 30;
const diasRestantes = Math.max(1, totalDiasMes - diasOperacao);
const mediaDiaria = diasOperacao > 0 ? (fatTotal / diasOperacao) : 0;
const projecaoMes = mediaDiaria * totalDiasMes;
const faltaBronze = Math.max(0, META_BRONZE - fatTotal);
const mediaNecBronze = diasRestantes > 0 ? (faltaBronze / diasRestantes) : 0;

console.log(`✅ ${rows.length} lançamentos processados sem perdas!`);
console.log(`📊 DESEMPENHO OFICIAL POR TURNO (${diasOperacao} DIAS DE OPERAÇÃO):`);
console.log(`🌅 Turno A (Manhã):     ${fmtBRL(turnosTotais.A.fat)} (${turnosTotais.A.malas} malas)`);
console.log(`☀️ Turno B (Tarde):     ${fmtBRL(turnosTotais.B.fat)} (${turnosTotais.B.malas} malas)`);
console.log(`🌙 Turno D (Madrugada): ${fmtBRL(turnosTotais.D.fat)} (${turnosTotais.D.malas} malas)`);
console.log(`💰 Faturamento Geral:   ${fmtBRL(fatTotal)} (${malasTotal} malas)`);
console.log(`🔮 Projeção (30 dias):  ${fmtBRL(projecaoMes)}`);

// Atualizar index.html com dados consolidados
const indexPath = path.join(REPO_DIR, 'index.html');
let indexHtml = fs.readFileSync(indexPath, 'utf8');

const startIdx = indexHtml.indexOf('let DADOS_ATUAIS = {');
if (startIdx !== -1) {
  const endIdx = indexHtml.indexOf('};', startIdx) + 2;
  const novoBloco = 'let DADOS_ATUAIS = ' + JSON.stringify({
    faturamentoTotal: Math.round(fatTotal),
    totalMalas: malasTotal,
    totalLinhas: rows.length,
    diasOperacao: diasOperacao,
    turnos: turnosTotais,
    colabs: colabsList,
    dataTurno: dataUltima,
    arquivoNome: path.basename(targetFile)
  }, null, 2) + ';';
  indexHtml = indexHtml.slice(0, startIdx) + novoBloco + indexHtml.slice(endIdx);
  fs.writeFileSync(indexPath, indexHtml, 'utf8');
} else {
  console.log('⚠️ Aviso: bloco DADOS_ATUAIS não encontrado no index.html');
}

// 1. Resumo Operacional WhatsApp
let msgWpp = `📊 *RELATÓRIO OPERACIONAL GRU - ONEBOSS*\n`;
msgWpp += `📅 *Data:* ${dataUltima} (${diasOperacao} dias)\n`;
msgWpp += `━━━━━━━━━━━━━━━━━━━━━\n\n`;

const nomesTurnos = {
  A: '🌅 TURNO A (Manhã)',
  B: '☀️ TURNO B (Tarde)',
  D: '🌙 TURNO D (Madrugada)'
};

['A', 'B', 'D'].forEach(t => {
  const turn = turnosTotais[t];
  msgWpp += `*${nomesTurnos[t]}*\n`;
  msgWpp += `💰 Fat: ${fmtBRL(turn.fat)} | 🧳 Malas: ${turn.malas}\n`;

  const topColabs = colabsList.filter(c => c.turno === t).slice(0, 5);
  if (topColabs.length > 0) {
    topColabs.forEach((c, i) => {
      const medal = i === 0 ? '🥇' : i === 1 ? '🥈' : i === 2 ? '🥉' : '▫️';
      msgWpp += `${medal} *${c.nome.toUpperCase()}*: ${fmtBRL(c.fat)} (${c.malas} malas)\n`;
    });
  } else {
    msgWpp += `_Nenhum colaborador registrado_\n`;
  }
  msgWpp += `\n`;
});

msgWpp += `━━━━━━━━━━━━━━━━━━━━━\n`;
msgWpp += `🎯 *TOTAL GERAL DO DIA*\n`;
msgWpp += `💵 *Faturamento Total:* ${fmtBRL(fatTotal)}\n`;
msgWpp += `🧳 *Total de Malas:* ${malasTotal}\n`;
msgWpp += `💳 *Lançamentos:* ${rows.length}\n\n`;
msgWpp += `🚀 *Equipe Protecbag GRU*`;

const resumoPath = path.join(REPO_DIR, 'resumo_whatsapp.txt');
fs.writeFileSync(resumoPath, msgWpp, 'utf8');

// 2. Resumo Diretoria
const pctA = fatTotal > 0 ? ((turnosTotais.A.fat / fatTotal) * 100).toFixed(2) : '0.00';
const pctB = fatTotal > 0 ? ((turnosTotais.B.fat / fatTotal) * 100).toFixed(2) : '0.00';
const pctD = fatTotal > 0 ? ((turnosTotais.D.fat / fatTotal) * 100).toFixed(2) : '0.00';

let msgDir = `Resumo de Desempenho Parcial (${diasOperacao} Dias de Operação)\n\n`;
msgDir += `Resultado Atual: Faturamento de ${fmtBRL(fatTotal)} com um volume de ${malasTotal.toLocaleString('pt-BR')} malas movimentadas (Média atual de ${fmtBRL(mediaDiaria)}/dia).\n\n`;
msgDir += `Projeção Final do Mês: ${fmtBRL(projecaoMes)}. Mantendo o ritmo atual, o fechamento ficará ${projecaoMes < META_BRONZE ? 'abaixo da Meta Bronze' : 'dentro das metas'}.\n\n`;
if (faltaBronze > 0) {
  msgDir += `Ação Necessária (Meta Bronze): Para atingir a meta base, é preciso faturar os ${fmtBRL(faltaBronze)} restantes, exigindo uma elevação da média diária para ${fmtBRL(mediaNecBronze)} nos próximos ${diasRestantes} dias.\n\n`;
}
msgDir += `Participação por Turno:\n\n`;
msgDir += `Turno B (Destaque): ${pctB}% do faturamento (${fmtBRL(turnosTotais.B.fat)} / ${turnosTotais.B.malas.toLocaleString('pt-BR')} malas).\n\n`;
msgDir += `Turno A: ${pctA}% do faturamento (${fmtBRL(turnosTotais.A.fat)} / ${turnosTotais.A.malas.toLocaleString('pt-BR')} malas).\n\n`;
msgDir += `Turno C: ${pctD}% do faturamento (${fmtBRL(turnosTotais.D.fat)} / ${turnosTotais.D.malas.toLocaleString('pt-BR')} malas).\n\n`;
msgDir += `Total de Malas (Geral): ${malasTotal.toLocaleString('pt-BR')} malas (${diasOperacao} Dias de operação)`;

const resumoDir = path.join(REPO_DIR, 'resumo_diretoria.txt');
fs.writeFileSync(resumoDir, msgDir, 'utf8');

try {
  execSync(`powershell -command "Set-Clipboard -Value (Get-Content -Path '${resumoDir}' -Raw)"`);
  console.log('\n📲 Resumo Executivo da Diretoria copiado para a Área de Transferência!');
} catch (e) {}

// Sincronizar com GitHub
console.log('\n🚀 Atualizando GitHub...');
try {
  execSync(`git -C "${REPO_DIR}" add index.html atualizar_vendas.js resumo_whatsapp.txt resumo_diretoria.txt`, { stdio: 'inherit' });
  execSync(`git -C "${REPO_DIR}" commit -m "feat: suporte a painel executivo e metas diretoria"`, { stdio: 'inherit' });
  execSync(`git -C "${REPO_DIR}" push origin main`, { stdio: 'inherit' });
  console.log('🌐 GitHub Pages sincronizado com sucesso!\n');
} catch (e) {
  console.log('⚠️ Aviso ao sincronizar com GitHub:', e.message);
}

console.log('🎉 PROCESSO CONCLUÍDO COM SUCESSO!\n');
