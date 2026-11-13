// one-boss-report.js
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

// ----- CONFIG ---------------------------------------------------------
const UPLOADS_DIR = path.resolve(__dirname);          // pasta onde o .xlsx chega
const ONEBOSS_WEBHOOK = process.env.ONEBOSS_WEBHOOK; // opcional webhook
// ---------------------------------------------------------------------

// encontra o último .xlsx (padrão usado pelo portal)
function getLatestXlsx() {
  const files = fs.readdirSync(UPLOADS_DIR)
    .filter(f => f.endsWith('.xlsx'))
    .map(f => ({ name: f, time: fs.statSync(path.join(UPLOADS_DIR, f)).mtimeMs }))
    .sort((a, b) => b.time - a.time);
  if (!files.length) throw new Error('Nenhum .xlsx encontrado');
  return path.join(UPLOADS_DIR, files[0].name);
}

// inferência de turno OneBoss precisa ser idêntica ao index.html
function inferirTurnoOneBoss(row) {
  const turnoId = String(row['Turno ID'] || row['TurnoID'] || '').trim();
  const nomeColab = String(row['Funcionário'] || row['Funcionario'] || row['Nome'] || '').toLowerCase();
  const turnoNum = String(row['Turno'] || '').trim();

  // 1️⃣ Pelo Turno ID de fechamento
  if (turnoId === '001') return 'A';
  if (turnoId === '002') return 'B';
  if (turnoId === '003' || turnoId === '004') return 'D';

  // 2️⃣ Por marcações no nome (ex: "d2", "manha", etc)
  if (nomeColab.endsWith('d2') || nomeColab.includes('madrugada')) return 'D';
  if (nomeColab.includes('manha')) return 'A';
  if (nomeColab.includes('tarde')) return 'B';

  // 3️⃣ Pelo número do Turno na linha
  if (turnoNum === '1') return 'A';
  if (turnoNum === '2') return 'B';
  if (turnoNum === '3' || turnoNum === '4') return 'D';

  // Fallback padrão
  return 'B';
}

function normalizarNome(nome) {
  if (!nome) return 'DESCONHECIDO';
  return nome.trim().toUpperCase();
}

function gerarRelatorio(dados) {
  const agora = new Date();
  const linhas = [];
  linhas.push(`📊 *RELATÓRIO OPERACIONAL GRU*`);
  linhas.push(`🕒 Atualizado em: ${agora.toLocaleDateString('pt-BR')} às ${agora.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' })}`);
  linhas.push(`━━━━━━━━━━━━━━━━━━━━━`);

  const nomesTurnos = {
    A: '🌅 TURNO A (Manhã)',
    B: '☀️ TURNO B (Tarde)',
    D: '🌙 TURNO D (Madrugada)'
  };

  for (const t of ['A', 'B', 'D']) {
    const turn = dados.turnos[t] || { faturamento: 0, malas: 0 };
    linhas.push(`\n*${nomesTurnos[t]}*`);
    linhas.push(`💰 Fat: R$ ${turn.faturamento.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} | 🧳 Malas: ${turn.malas}`);
    
    // Lista top do turno
    const colabs = Object.values(dados.colabsMap)
      .filter(c => c.turno === t)
      .sort((a, b) => b.faturamento - a.faturamento);

    if (colabs.length > 0) {
      colabs.slice(0, 5).forEach((c, i) => {
        const medalha = i === 0 ? '🥇' : i === 1 ? '🥈' : i === 2 ? '🥉' : '▫️';
        linhas.push(`${medalha} *${c.nome}*: R$ ${c.faturamento.toLocaleString('pt-BR', { minimumFractionDigits: 0, maximumFractionDigits: 0 })} (${c.malas} malas)`);
      });
    } else {
      linhas.push(`_Nenhuma venda registrada ainda_`);
    }
  }

  const totalFatGeral = dados.turnos.A.faturamento + dados.turnos.B.faturamento + dados.turnos.D.faturamento;
  const totalMalasGeral = dados.turnos.A.malas + dados.turnos.B.malas + dados.turnos.D.malas;

  linhas.push(`\n━━━━━━━━━━━━━━━━━━━━━`);
  linhas.push(`🎯 *TOTAL GERAL DO DIA*`);
  linhas.push(`💵 Total: R$ ${totalFatGeral.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`);
  linhas.push(`🧳 Malas Total: ${totalMalasGeral}`);

  return linhas.join('\n');
}

// ---- MAIN ------------------------------------------------------------
async function executar() {
  try {
    const xlsxPath = getLatestXlsx();
    console.log('📂 Lendo arquivo:', path.basename(xlsxPath));
    const wb = XLSX.readFile(xlsxPath);
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });

    const DADOS = {
      turnos: {
        A: { faturamento: 0, malas: 0 },
        B: { faturamento: 0, malas: 0 },
        D: { faturamento: 0, malas: 0 }
      },
      colabsMap: {}
    };

    for (const r of rows) {
      const nome = normalizarNome(r['Funcionário'] || r['Funcionario'] || r['Nome'] || r['COLABORADOR'] || '');
      const valor = parseFloat(r['Total Venda'] || r['Valor'] || r['Faturamento'] || 0) || 0;
      const turno = inferirTurnoOneBoss(r);

      // Soma por turno
      DADOS.turnos[turno].faturamento += valor;
      DADOS.turnos[turno].malas += 1;

      // Agrupa colaborador
      if (!DADOS.colabsMap[nome]) {
        DADOS.colabsMap[nome] = { nome, turno, faturamento: 0, malas: 0 };
      }
      DADOS.colabsMap[nome].faturamento += valor;
      DADOS.colabsMap[nome].malas += 1;
    }

    const texto = gerarRelatorio(DADOS);

    // Se webhook OneBoss configurado
    if (ONEBOSS_WEBHOOK) {
      await fetch(ONEBOSS_WEBHOOK, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ texto, horario: new Date().toISOString() })
      });
      console.log('📡 Relatório enviado para Webhook OneBoss com sucesso!');
    }

    // Link WhatsApp
    const whatsappUrl = `https://api.whatsapp.com/send?text=${encodeURIComponent(texto)}`;
    console.log('\n💬 Relatório formatado para WhatsApp:\n');
    console.log(texto);
    console.log('\n🔗 Link direto WhatsApp:\n', whatsappUrl);

    // Salva em arquivo
    const outPath = path.join(__dirname, 'ultimo_relatorio_hora.txt');
    fs.writeFileSync(outPath, texto, 'utf8');
    console.log('\n💾 Salvo em:', outPath);
    return texto;
  } catch (e) {
    console.error('❌ Erro ao gerar relatório:', e);
  }
}

if (require.main === module) {
  executar();
}

module.exports = { executar };
