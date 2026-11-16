"""
=============================================================================
  IMPORTADOR OFICIAL DA ESCALA DE DEZEMBRO (PRODUÇÃO GRU)
  Facilita o upload manual da escala oficial enviada pelo Gerente.
=============================================================================
"""

import os
import sys
import re
import json
import subprocess
import glob
import openpyxl

CAMINHO_HTML = os.path.join(os.path.dirname(__file__), "index.html")

def normalizar_codigo_dia(val):
    if val is None:
        return "F"
    s = str(val).strip().upper()
    if s in ["1", "1.0", "T", "TRAB", "TRABALHO", "P", "PRESENCA", "X", "SIM", "OK"]:
        return "1"
    if s in ["F", "FOLGA", "DSR", "DF", "0", "0.0", "FERIAS", "FER", "ATESTADO", "LICENCA"]:
        return "F"
    if s == "1":
        return "1"
    if "F" in s:
        return "F"
    return "1"

def localizar_arquivo_escala():
    # 1. Se foi passado via linha de comando
    if len(sys.argv) > 1 and os.path.isfile(sys.argv[1]):
        return sys.argv[1]
    
    # 2. Procurar na pasta atual ou Desktop
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    padroes = [
        os.path.join(desktop, "*escala*dezembro*.xlsx"),
        os.path.join(desktop, "*escala*.xlsx"),
        os.path.join(os.path.dirname(__file__), "*escala*dezembro*.xlsx"),
        os.path.join(os.path.dirname(__file__), "*escala*.xlsx"),
        os.path.join(os.path.dirname(__file__), "modelo_escala_dezembro.xlsx"),
    ]
    
    arquivos_encontrados = []
    for p in padroes:
        for f in glob.glob(p):
            if f not in arquivos_encontrados and os.path.isfile(f):
                arquivos_encontrados.append(f)
                
    if arquivos_encontrados:
        print("\n🔍 Arquivos de escala encontrados no sistema:")
        for idx, f in enumerate(arquivos_encontrados, start=1):
            nome_arq = os.path.basename(f)
            pasta = "Desktop" if desktop in f else "escala-gru"
            print(f"  [{idx}] {nome_arq} ({pasta})")
        print("  [0] Digitar ou arrastar outro arquivo manualmente")
        
        escolha = input("\nEscolha o número do arquivo (ou 0 para manual): ").strip()
        if escolha.isdigit() and 1 <= int(escolha) <= len(arquivos_encontrados):
            return arquivos_encontrados[int(escolha) - 1]
            
    # Entrada manual
    print("\n📂 Por favor, arraste o arquivo Excel (.xlsx) da escala aqui ou digite o caminho:")
    caminho = input("Caminho do arquivo: ").strip(' "\'')
    if os.path.isfile(caminho):
        return caminho
        
    return None

def processar_escala_excel(caminho_arquivo):
    print(f"\n📖 Lendo arquivo: {caminho_arquivo}")
    wb = openpyxl.load_workbook(caminho_arquivo, data_only=True)
    sheet = wb.active
    print(f"📋 Aba ativa: '{sheet.title}' (Linhas: {sheet.max_row}, Colunas: {sheet.max_column})")
    
    # Identificar mapeamento de colunas para dias 1 a 31
    colunas_dias = {}
    for r in range(1, min(10, sheet.max_row + 1)):
        dias_detectados = {}
        for c in range(2, min(40, sheet.max_column + 1)):
            val = sheet.cell(r, c).value
            if val is not None:
                s_val = str(val).strip()
                if s_val.isdigit() and 1 <= int(s_val) <= 31:
                    dias_detectados[int(s_val)] = c
        if len(dias_detectados) >= 15:
            colunas_dias = dias_detectados
            break

    if not colunas_dias:
        # Fallback padrão: colunas 2 a 32 representam dias 1 a 31
        colunas_dias = {d: d + 1 for d in range(1, 32)}

    # Mapear seções na planilha (TURNO A, TURNO B, LIDER, etc.)
    secoes = {}
    secao_atual = "TURNO A"
    secoes[secao_atual] = []
    
    for r in range(1, sheet.max_row + 1):
        val_col1 = sheet.cell(r, 1).value
        if not val_col1:
            continue
        txt = str(val_col1).strip().upper()
        
        # Mudança de seção
        if "TURNO A" in txt:
            secao_atual = "TURNO A"
            if secao_atual not in secoes: secoes[secao_atual] = []
            continue
        elif "TURNO B" in txt:
            secao_atual = "TURNO B"
            if secao_atual not in secoes: secoes[secao_atual] = []
            continue
        elif "LIDER" in txt:
            secao_atual = "LIDERES"
            if secao_atual not in secoes: secoes[secao_atual] = []
            continue
        elif any(k in txt for k in ["TURNO C", "TURNO D", "OUTROS", "REFORCO"]):
            secao_atual = txt
            if secao_atual not in secoes: secoes[secao_atual] = []
            continue
        elif any(k in txt for k in ["QTA.", "TOTAL", "NOVEMBRO", "DEZEMBRO", "JANEIRO", "COLABORADOR", "DOM", "SEG", "TER"]):
            continue
            
        # Nome de colaborador válido (> 2 caracteres)
        if len(txt) > 2:
            dias = []
            for d in range(1, 32):
                col_idx = colunas_dias.get(d)
                if col_idx:
                    raw_val = sheet.cell(r, col_idx).value
                    dias.append(normalizar_codigo_dia(raw_val))
                else:
                    dias.append("1")
                    
            secoes[secao_atual].append({
                "nome": txt,
                "dias": dias
            })

    print(f"\n📑 Seções detectadas no arquivo:")
    for sec, colabs in secoes.items():
        if colabs:
            print(f"   • {sec}: {len(colabs)} pessoas")
            
    # O portal de escala principal utiliza o TURNO A
    alvo = secoes.get("TURNO A")
    if not alvo:
        # Se não houver divisão explícita, pega a primeira seção com colaboradores
        for sec, colabs in secoes.items():
            if colabs:
                alvo = colabs
                print(f"ℹ️ Usando colaboradores da seção: {sec}")
                break

    return alvo

def exibir_preview(colaboradores):
    print("\n" + "="*80)
    print("📊 PRÉVIA DA ESCALA DE DEZEMBRO - TURNO A (31 DIAS)")
    print("="*80)
    print(f"{'Nº':<3} | {'COLABORADOR':<26} | {'TRAB':<5} | {'FOLGAS':<6} | {'AMOSTRA (DIAS 1 A 15)':<25}")
    print("-"*80)
    
    for idx, c in enumerate(colaboradores, start=1):
        trab = c['dias'].count('1')
        folgas = c['dias'].count('F')
        amostra = " ".join(c['dias'][:15])
        print(f"{idx:<3} | {c['nome'][:26]:<26} | {trab:<5} | {folgas:<6} | {amostra}")
        
    print("-"*80)
    print(f"Total de Colaboradores no Turno A: {len(colaboradores)}")
    print("="*80 + "\n")

def atualizar_index_html(colaboradores):
    if not os.path.isfile(CAMINHO_HTML):
        print(f"❌ Erro: index.html não encontrado em {CAMINHO_HTML}")
        return False
        
    with open(CAMINHO_HTML, "r", encoding="utf-8") as f:
        html = f.read()
        
    # 1. Montar bloco JS TURNO_A_DEZ
    linhas_js = ["  const TURNO_A_DEZ = ["]
    for c in colaboradores:
        dias_json = json.dumps(c['dias'])
        linhas_js.append(f'    {{nome:"{c["nome"]}", dias:{dias_json}}},')
    linhas_js.append("  ];")
    novo_bloco_dez = "\n".join(linhas_js)
    
    # 2. Substituir TURNO_A_DEZ existente
    padrao_dez = r"const TURNO_A_DEZ = \[[\s\S]*?\];"
    if not re.search(padrao_dez, html):
        print("❌ Erro: Não foi encontrada a variável TURNO_A_DEZ em index.html")
        return False
        
    html_mod = re.sub(padrao_dez, novo_bloco_dez, html, count=1)
    
    # 3. Habilitar Dezembro no selector: remover 'disabled' e ajustar label
    html_mod = html_mod.replace(
        '<option value="2026-12" disabled>Dezembro 2026 (Aguardando Gerente)</option>',
        '<option value="2026-12">Dezembro 2026 (Oficial)</option>'
    )
    html_mod = html_mod.replace(
        '<option value="2026-12">Dezembro 2026</option>',
        '<option value="2026-12">Dezembro 2026 (Oficial)</option>'
    )
    
    with open(CAMINHO_HTML, "w", encoding="utf-8") as f:
        f.write(html_mod)
        
    print("✅ index.html atualizado com a escala oficial de Dezembro!")
    return True

def publicar_github():
    print("\n🌐 Publicando alterações no GitHub Pages...")
    repo_dir = os.path.dirname(CAMINHO_HTML)
    try:
        subprocess.run(["git", "add", "index.html"], cwd=repo_dir, check=True)
        subprocess.run(["git", "commit", "-m", "feat: publicar escala oficial de dezembro do gerente"], cwd=repo_dir, check=True)
        subprocess.run(["git", "push", "origin", "main"], cwd=repo_dir, check=True)
        print("🚀 SUCESSO! Escala publicada no GitHub Pages com sucesso!")
        print("🔗 Acesse em: https://gtsdiego94-cmd.github.io/escala-gru/")
        return True
    except Exception as e:
        print(f"⚠️ Aviso ao enviar para o GitHub: {e}")
        return False

def main():
    print("="*80)
    print("  🏭 ATUALIZADOR MANUAL DA ESCALA DE DEZEMBRO - GRU")
    print("="*80)
    
    caminho = localizar_arquivo_escala()
    if not caminho:
        print("❌ Nenhum arquivo selecionado. Operação cancelada.")
        input("\nPressione ENTER para sair...")
        return
        
    colaboradores = processar_escala_excel(caminho)
    if not colaboradores:
        print("❌ Nenhum colaborador encontrado no arquivo.")
        input("\nPressione ENTER para sair...")
        return
        
    exibir_preview(colaboradores)
    
    print("⚠️ ATENÇÃO:")
    print("Ao confirmar abaixo, a escala de Dezembro será ATIVADA no portal oficial da equipe.")
    print("Se ainda estiver aguardando aprovação final do gerente, responda 'N'.")
    
    resp = input("\nDeseja ATUALIZAR e PUBLICAR a escala de Dezembro agora? [S/N]: ").strip().upper()
    if resp in ["S", "SIM"]:
        if atualizar_index_html(colaboradores):
            publicar_github()
            print("\n🎉 Tudo pronto! Sua equipe agora já consegue visualizar Dezembro no portal!")
    else:
        print("\nℹ️ Operação cancelada pelo usuário.")
        print("A escala de Dezembro CONTINUA BLOQUEADA no portal aguardando seu comando.")
        
    input("\nPressione ENTER para finalizar...")

if __name__ == "__main__":
    main()
