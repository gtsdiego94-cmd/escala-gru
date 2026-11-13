import os
import glob
import json
import subprocess
import pandas as pd
from datetime import datetime

# Diretórios onde o OneBoss pode ter salvado o relatório
PASTAS_BUSCA = [
    r"C:\Relatorios",
    r"C:\Export",
    r"C:\Users\User\Desktop",
    r"C:\Users\User\Downloads",
    r"C:\Users\User\Desktop\escala-gru"
]

REPO_DIR = r"C:\Users\User\Desktop\escala-gru"
INDEX_PATH = os.path.join(REPO_DIR, "index.html")
RESUMO_PATH = os.path.join(REPO_DIR, "resumo_whatsapp.txt")

def inferir_turno(nome, turno_id, turno_num):
    tid = str(turno_id).strip() if pd.notna(turno_id) else ""
    tnum = str(turno_num).strip() if pd.notna(turno_num) else ""
    n = str(nome).strip().upper() if pd.notna(nome) else ""

    # 1. Pelo Turno ID do OneBoss
    if tid in ['001', '1']: return 'A'
    if tid in ['002', '2']: return 'B'
    if tid in ['003', '3', '004', '4']: return 'D'

    # 2. Pelo nome do colaborador
    if n.endswith('D2') or ' D2' in n or '(D2)' in n: return 'D'

    # 3. Pelo número simples de Turno
    if tnum == '1': return 'A'
    if tnum == '2': return 'B'
    if tnum in ['3', '4']: return 'D'

    return 'B'

def achar_ultimo_relatorio():
    candidatos = []
    for pasta in PASTAS_BUSCA:
        if os.path.exists(pasta):
            for padrao in ["*RelatorioMovimentoCaixa*.xlsx", "*RelatorioMovimentoCaixa*.csv", "*.xlsx", "*.csv"]:
                for arquivo in glob.glob(os.path.join(pasta, padrao)):
                    try:
                        mtime = os.path.getmtime(arquivo)
                        candidatos.append((arquivo, mtime))
                    except:
                        pass

    if not candidatos:
        return None

    candidatos.sort(key=lambda x: x[1], reverse=True)
    return candidatos[0][0]

def fmt_brl(v):
    return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def processar_relatorio():
    print("=" * 60)
    print("  🚀 ATUALIZADOR ONEBOSS EM PYTHON (GRU)")
    print("=" * 60)

    arquivo = achar_ultimo_relatorio()
    if not arquivo:
        print("\n❌ Nenhum relatório encontrado nas pastas configuradas!")
        return

    print(f"\n📂 Arquivo mais recente detectado:\n   --> {arquivo}")

    # Leitura com Pandas
    if arquivo.endswith(".xlsx") or arquivo.endswith(".xls"):
        df = pd.read_excel(arquivo)
    else:
        df = pd.read_csv(arquivo, encoding="latin1", sep=None, engine="python")

    total_linhas = len(df)
    print(f"📊 Total de registros lidos: {total_linhas}")

    # Identifica colunas
    col_valor = "Valor" if "Valor" in df.columns else "Total Venda"
    col_func = "Funcionário" if "Funcionário" in df.columns else ("Funcionario" if "Funcionario" in df.columns else "Nome")
    col_tipo = "Tipo" if "Tipo" in df.columns else ""
    col_item = "Item" if "Item" in df.columns else ""
    col_tid = "Turno ID" if "Turno ID" in df.columns else "TurnoID"
    col_tnum = "Turno" if "Turno" in df.columns else ""
    col_data = "Data" if "Data" in df.columns else ""

    fat_total = 0.0
    malas_total = 0
    turnos = {
        'A': {'fat': 0.0, 'malas': 0},
        'B': {'fat': 0.0, 'malas': 0},
        'D': {'fat': 0.0, 'malas': 0}
    }
    colabs_map = {}
    data_ultima = datetime.now().strftime("%d/%m/%Y")

    for _, row in df.iterrows():
        func_nome = str(row[col_func]).strip() if pd.notna(row[col_func]) else ""
        if not func_nome or func_nome.lower() == "total":
            continue

        raw_val = row[col_valor]
        try:
            if isinstance(raw_val, str):
                raw_val = raw_val.replace("R$", "").replace(".", "").replace(",", ".").strip()
            val = float(raw_val)
        except:
            val = 0.0

        tipo_str = str(row[col_tipo]).lower() if col_tipo and pd.notna(row[col_tipo]) else ""
        item_str = str(row[col_item]).lower() if col_item and pd.notna(row[col_item]) else ""
        is_mala = "servi" in tipo_str or "aerobag" in item_str or "lacre" in item_str or "reembalar" in item_str

        tid = row[col_tid] if col_tid in row else ""
        tnum = row[col_tnum] if col_tnum in row else ""
        turno = inferir_turno(func_nome, tid, tnum)

        fat_total += val
        if is_mala: malas_total += 1

        turnos[turno]['fat'] += val
        if is_mala: turnos[turno]['malas'] += 1

        k = func_nome.upper()
        if k not in colabs_map:
            colabs_map[k] = {'nome': func_nome, 'turno': turno, 'fat': 0.0, 'malas': 0}
        colabs_map[k]['fat'] += val
        if is_mala: colabs_map[k]['malas'] += 1

    colabs_list = []
    for c in sorted(colabs_map.values(), key=lambda x: x['fat'], reverse=True):
        media = round(c['fat'] / c['malas'], 2) if c['malas'] > 0 else 0.0
        colabs_list.append({
            'pos': len(colabs_list) + 1,
            'nome': c['nome'],
            'turno': c['turno'],
            'fat': round(c['fat']),
            'malas': c['malas'],
            'media': media
        })

    print("\n" + "─" * 40)
    print("📈 RESULTADO CALCULADO EM PYTHON:")
    print(f"🌅 Turno A (Manhã):     {fmt_brl(turnos['A']['fat'])} ({turnos['A']['malas']} malas)")
    print(f"☀️ Turno B (Tarde):     {fmt_brl(turnos['B']['fat'])} ({turnos['B']['malas']} malas)")
    print(f"🌙 Turno D (Madrugada): {fmt_brl(turnos['D']['fat'])} ({turnos['D']['malas']} malas)")
    print(f"💰 Faturamento Geral:   {fmt_brl(fat_total)} ({malas_total} malas)")
    print("─" * 40)

    # Injeção no index.html
    if os.path.exists(INDEX_PATH):
        with open(INDEX_PATH, "r", encoding="utf-8") as f:
            html = f.read()

        start = html.find("let DADOS_ATUAIS = {")
        if start != -1:
            end = html.find("};", start) + 2
            novo_dados = {
                "faturamentoTotal": round(fat_total),
                "totalMalas": malas_total,
                "totalLinhas": total_linhas,
                "turnos": turnos,
                "colabs": colabs_list,
                "dataTurno": data_ultima,
                "arquivoNome": os.path.basename(arquivo)
            }
            bloco_subst = "let DADOS_ATUAIS = " + json.dumps(novo_dados, indent=2, ensure_ascii=False) + ";"
            html = html[:start] + bloco_subst + html[end:]
            with open(INDEX_PATH, "w", encoding="utf-8") as f:
                f.write(html)
            print("✅ index.html atualizado com sucesso!")
        else:
            print("⚠️ Aviso: bloco DADOS_ATUAIS não localizado no index.html")

    # Resumo para WhatsApp
    msg_wpp = "📊 *RELATÓRIO OPERACIONAL GRU - ONEBOSS*\n"
    msg_wpp += f"📅 *Data:* {data_ultima}\n"
    msg_wpp += "━━━━━━━━━━━━━━━━━━━━━\n\n"

    nomes_t = {'A': '🌅 TURNO A (Manhã)', 'B': '☀️ TURNO B (Tarde)', 'D': '🌙 TURNO D (Madrugada)'}
    for t in ['A', 'B', 'D']:
        turn = turnos[t]
        msg_wpp += f"*{nomes_t[t]}*\n"
        msg_wpp += f"💰 Fat: {fmt_brl(turn['fat'])} | 🧳 Malas: {turn['malas']}\n"
        top_c = [c for c in colabs_list if c['turno'] == t][:5]
        if top_c:
            for i, c in enumerate(top_c):
                medal = "🥇" if i == 0 else ("🥈" if i == 1 else ("🥉" if i == 2 else "▫️"))
                msg_wpp += f"{medal} *{c['nome'].upper()}*: {fmt_brl(c['fat'])} ({c['malas']} malas)\n"
        else:
            msg_wpp += "_Nenhum colaborador registrado_\n"
        msg_wpp += "\n"

    msg_wpp += "━━━━━━━━━━━━━━━━━━━━━\n"
    msg_wpp += "🎯 *TOTAL GERAL DO DIA*\n"
    msg_wpp += f"💵 *Faturamento Total:* {fmt_brl(fat_total)}\n"
    msg_wpp += f"🧳 *Total de Malas:* {malas_total}\n"
    msg_wpp += f"💳 *Lançamentos:* {total_linhas}\n\n"
    msg_wpp += "🚀 *Equipe Protecbag GRU*"

    with open(RESUMO_PATH, "w", encoding="utf-8") as f:
        f.write(msg_wpp)

    # Copia para área de transferência via powershell
    try:
        subprocess.run(["powershell", "-command", f"Set-Clipboard -Value (Get-Content -Path '{RESUMO_PATH}' -Raw)"], check=False)
        print("📲 Resumo WhatsApp copiado para a Área de Transferência!")
    except:
        pass

    # Sincroniza com GitHub
    print("\n🌐 Enviando para o GitHub Pages...")
    try:
        subprocess.run(["git", "-C", REPO_DIR, "add", "index.html", "resumo_whatsapp.txt"], check=False)
        subprocess.run(["git", "-C", REPO_DIR, "commit", "-m", f"update: atualizacao via leitor python ({data_ultima})"], check=False)
        subprocess.run(["git", "-C", REPO_DIR, "push", "origin", "main"], check=False)
        print("🎉 Página online atualizada no GitHub!")
    except Exception as e:
        print(f"⚠️ Erro ao enviar para o Git: {e}")

if __name__ == "__main__":
    processar_relatorio()
