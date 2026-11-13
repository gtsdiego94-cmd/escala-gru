import os
import sys
import glob
import json
import subprocess
import pandas as pd
from datetime import datetime
from gerar_card_imagem import gerar_card_imagem

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
RESUMO_DIRETORIA_PATH = os.path.join(REPO_DIR, "resumo_diretoria.txt")

META_BRONZE = 1500000.0
META_PRATA  = 1600000.0
META_OURO   = 1650000.0

def inferir_turno(nome, turno_id, turno_num):
    tid = str(turno_id).strip() if pd.notna(turno_id) else ""
    tnum = str(turno_num).strip() if pd.notna(turno_num) else ""
    n = str(nome).strip().upper() if pd.notna(nome) else ""

    # 1. Pelo Turno ID do OneBoss (001=A, 002=B, 003=D, 004=D)
    if tid in ['001', '1']: return 'A'
    if tid in ['002', '2']: return 'B'
    if tid in ['003', '3', '004', '4']: return 'D'

    # 2. Pelo nome do colaborador (sufixo D2 = Turno D / Madrugada)
    if n.endswith('D2') or ' D2' in n or '(D2)' in n: return 'D'

    # 3. Pelo número simples de Turno
    if tnum == '1': return 'A'
    if tnum == '2': return 'B'
    if tnum in ['3', '4']: return 'D'

    return 'B'

def achar_ultimo_relatorio():
    if len(sys.argv) > 1 and os.path.exists(sys.argv[1]):
        return sys.argv[1]

    candidatos = []
    for pasta in PASTAS_BUSCA:
        if os.path.exists(pasta):
            for padrao in ["*RelatorioMovimentoCaixa*.xlsx", "*RelatorioMovimentoCaixa*.csv", "*.xlsx", "*.csv"]:
                for arquivo in glob.glob(os.path.join(pasta, padrao)):
                    try:
                        base = os.path.basename(arquivo)
                        if base.startswith("~$") or "vendas_atual" in base:
                            continue
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
    print("=" * 65)
    print("  🚀 ATUALIZADOR ONEBOSS EM PYTHON (GRU) - DIRETORIA & METAS")
    print("=" * 65)

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

    fat_total = 0.0
    malas_total = 0
    turnos = {
        'A': {'fat': 0.0, 'malas': 0},
        'B': {'fat': 0.0, 'malas': 0},
        'D': {'fat': 0.0, 'malas': 0}
    }
    colabs_map = {}
    data_ultima = datetime.now().strftime("%d/%m/%Y")

    # Contagem dinâmica de dias de operação
    dias_unicos = set()
    for col_dt in ["Data do Turno", "Data", "DATA"]:
        if col_dt in df.columns:
            try:
                s = pd.to_datetime(df[col_dt], errors='coerce').dt.date.dropna()
                if len(s) > 0:
                    dias_unicos = set(s)
                    break
            except:
                pass

    dias_operacao = len(dias_unicos) if len(dias_unicos) > 0 else datetime.now().day

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

    # Métricas Gerenciais & Diretoria
    total_dias_mes = 30
    dias_restantes = max(1, total_dias_mes - dias_operacao)
    media_diaria = fat_total / dias_operacao if dias_operacao > 0 else 0.0
    projecao_mes = media_diaria * total_dias_mes

    falta_bronze = max(0.0, META_BRONZE - fat_total)
    media_nec_bronze = falta_bronze / dias_restantes if dias_restantes > 0 else 0.0

    falta_prata = max(0.0, META_PRATA - fat_total)
    media_nec_prata = falta_prata / dias_restantes if dias_restantes > 0 else 0.0

    falta_ouro = max(0.0, META_OURO - fat_total)
    media_nec_ouro = falta_ouro / dias_restantes if dias_restantes > 0 else 0.0

    pct_a = (turnos['A']['fat'] / fat_total) * 100 if fat_total > 0 else 0.0
    pct_b = (turnos['B']['fat'] / fat_total) * 100 if fat_total > 0 else 0.0
    pct_d = (turnos['D']['fat'] / fat_total) * 100 if fat_total > 0 else 0.0

    print("\n" + "─" * 45)
    print(f"📈 RESULTADO CONSOLIDADO ({dias_operacao} DIAS DE OPERAÇÃO):")
    print(f"🌅 Turno A (Manhã):     {fmt_brl(turnos['A']['fat'])} ({turnos['A']['malas']} malas) - {pct_a:.2f}%")
    print(f"☀️ Turno B (Tarde):     {fmt_brl(turnos['B']['fat'])} ({turnos['B']['malas']} malas) - {pct_b:.2f}%")
    print(f"🌙 Turno D/C (Madrug.): {fmt_brl(turnos['D']['fat'])} ({turnos['D']['malas']} malas) - {pct_d:.2f}%")
    print(f"💰 Faturamento Geral:   {fmt_brl(fat_total)} ({malas_total} malas)")
    print(f"📊 Média Diária:        {fmt_brl(media_diaria)}/dia")
    print(f"🔮 Projeção (30 dias):  {fmt_brl(projecao_mes)}")
    print(f"🎯 Falta Meta Bronze:   {fmt_brl(falta_bronze)} (Ritmo: {fmt_brl(media_nec_bronze)}/dia)")
    print("─" * 45)

    # Injeção cirúrgica no index.html
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
                "diasOperacao": dias_operacao,
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

    # 1. Resumo Operacional para WhatsApp (Turnos & Top Colaboradores)
    msg_wpp = "📊 *RELATÓRIO OPERACIONAL GRU - ONEBOSS*\n"
    msg_wpp += f"📅 *Data:* {data_ultima} ({dias_operacao} dias de operação)\n"
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
    msg_wpp += "🎯 *TOTAL GERAL ACUMULADO*\n"
    msg_wpp += f"💵 *Faturamento Total:* {fmt_brl(fat_total)}\n"
    msg_wpp += f"🧳 *Total de Malas:* {malas_total}\n"
    msg_wpp += f"💳 *Lançamentos:* {total_linhas}\n\n"
    msg_wpp += "🚀 *Equipe Protecbag GRU*"

    with open(RESUMO_PATH, "w", encoding="utf-8") as f:
        f.write(msg_wpp)

    # 2. Resumo Executivo para a Diretoria (Metas & Projeção)
    status_proj = ""
    if projecao_mes >= META_OURO:
        status_proj = "🏆 Superando a Meta Ouro!"
    elif projecao_mes >= META_PRATA:
        status_proj = "🥈 Superando a Meta Prata!"
    elif projecao_mes >= META_BRONZE:
        status_proj = "🥉 Dentro da Meta Bronze (Base)."
    else:
        gap = META_BRONZE - projecao_mes
        status_proj = f"Mantendo o ritmo atual, o fechamento ficará abaixo da Meta Bronze (R$ 1,5M) por {fmt_brl(gap)}."

    def fmt_num(n):
        return f"{n:,}".replace(",", ".")

    def fmt_pct(p):
        return f"{p:.2f}".replace(".", ",")

    msg_dir = f"Resumo de Desempenho Parcial ({dias_operacao} Dias de Operação)\n\n"
    msg_dir += f"Resultado Atual: Faturamento de {fmt_brl(fat_total)} com um volume de {fmt_num(malas_total)} malas movimentadas (Média atual de {fmt_brl(media_diaria)}/dia).\n\n"
    msg_dir += f"Projeção Final do Mês: {fmt_brl(projecao_mes)}. {status_proj}\n\n"
    
    if falta_bronze > 0:
        msg_dir += f"Ação Necessária (Meta Bronze): Para atingir a meta base, é preciso faturar os {fmt_brl(falta_bronze)} restantes, exigindo uma elevação da média diária para {fmt_brl(media_nec_bronze)} nos próximos {dias_restantes} dias.\n\n"
    else:
        msg_dir += f"Meta Bronze atingida com sucesso! Rumo à Meta Prata ({fmt_brl(META_PRATA)}).\n\n"

    msg_dir += "Participação por Turno:\n\n"
    media_b = turnos['B']['fat'] / dias_operacao if dias_operacao > 0 else 0.0
    media_a = turnos['A']['fat'] / dias_operacao if dias_operacao > 0 else 0.0
    media_d = turnos['D']['fat'] / dias_operacao if dias_operacao > 0 else 0.0

    msg_dir += f"Turno B (Destaque): {fmt_pct(pct_b)}% do faturamento ({fmt_brl(turnos['B']['fat'])} / {fmt_num(turnos['B']['malas'])} malas) • Média: {fmt_brl(media_b)}/dia.\n\n"
    msg_dir += f"Turno A: {fmt_pct(pct_a)}% do faturamento ({fmt_brl(turnos['A']['fat'])} / {fmt_num(turnos['A']['malas'])} malas) • Média: {fmt_brl(media_a)}/dia.\n\n"
    msg_dir += f"Turno C: {fmt_pct(pct_d)}% do faturamento ({fmt_brl(turnos['D']['fat'])} / {fmt_num(turnos['D']['malas'])} malas) • Média: {fmt_brl(media_d)}/dia.\n\n"
    msg_dir += f"Total de Malas (Geral): {fmt_num(malas_total)} malas ({dias_operacao} Dias de operação)"

    with open(RESUMO_DIRETORIA_PATH, "w", encoding="utf-8") as f:
        f.write(msg_dir)

    # 3. GERAÇÃO AUTOMÁTICA DO CARD EXECUTIVO EM IMAGEM PNG
    try:
        pct_b_total = (turnos['B']['fat'] / META_BRONZE) * 100 if META_BRONZE else 0
        pct_p_total = (turnos['B']['fat'] / META_PRATA) * 100 if META_PRATA else 0
        pct_o_total = (turnos['B']['fat'] / META_OURO) * 100 if META_OURO else 0

        dados_card = {
            'periodo': f"Novembro / {datetime.now().year}",
            'dias_op': dias_operacao,
            'dias_rest': dias_restantes,
            'fat_total': fmt_brl(fat_total),
            'total_malas': fmt_num(malas_total),
            'media_diaria': fmt_brl(media_diaria),
            'projecao': fmt_brl(projecao_mes),
            'status_proj': status_proj,
            'falta_bronze': fmt_brl(falta_bronze),
            'media_nec_bronze': fmt_brl(media_nec_bronze),
            'pct_bronze': fmt_pct((fat_total / META_BRONZE) * 100) + '%',
            'falta_prata': fmt_brl(falta_prata),
            'media_nec_prata': fmt_brl(media_nec_prata),
            'pct_prata': fmt_pct((fat_total / META_PRATA) * 100) + '%',
            'falta_ouro': fmt_brl(falta_ouro),
            'media_nec_ouro': fmt_brl(media_nec_ouro),
            'pct_ouro': fmt_pct((fat_total / META_OURO) * 100) + '%',
            'fat_b': fmt_brl(turnos['B']['fat']),
            'malas_b': fmt_num(turnos['B']['malas']),
            'pct_b': fmt_pct(pct_b),
            'media_b': fmt_brl(media_b),
            'fat_a': fmt_brl(turnos['A']['fat']),
            'malas_a': fmt_num(turnos['A']['malas']),
            'pct_a': fmt_pct(pct_a),
            'media_a': fmt_brl(media_a),
            'fat_d': fmt_brl(turnos['D']['fat']),
            'malas_d': fmt_num(turnos['D']['malas']),
            'pct_d': fmt_pct(pct_d),
            'media_d': fmt_brl(media_d),
            'data_hora': datetime.now().strftime("%d/%m/%Y %H:%M")
        }

        desktop_img = os.path.join(r"C:\Users\User\Desktop", "Relatorio_Diretoria_GRU.png")
        repo_img = os.path.join(REPO_DIR, "relatorio_diretoria.png")
        gerar_card_imagem(dados_card, desktop_img)
        gerar_card_imagem(dados_card, repo_img)
    except Exception as err_img:
        print(f"⚠️ Erro ao gerar card em imagem: {err_img}")

    # Sincroniza com GitHub
    print("\n🌐 Enviando para o GitHub Pages...")
    try:
        subprocess.run(["git", "-C", REPO_DIR, "add", "index.html", "resumo_whatsapp.txt", "resumo_diretoria.txt", "relatorio_diretoria.png", "gerar_card_imagem.py", "atualizar_vendas_python.py"], check=False)
        subprocess.run(["git", "-C", REPO_DIR, "commit", "-m", f"feat: geracao automatica de imagem executiva para diretoria ({data_ultima})"], check=False)
        subprocess.run(["git", "-C", REPO_DIR, "push", "origin", "main"], check=False)
        print("🎉 Página online atualizada no GitHub!")
    except Exception as e:
        print(f"⚠️ Erro ao enviar para o Git: {e}")

if __name__ == "__main__":
    processar_relatorio()
