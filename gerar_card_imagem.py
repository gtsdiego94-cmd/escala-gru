import os
import subprocess
from PIL import Image, ImageDraw, ImageFont

def gerar_card_imagem(dados, caminho_saida):
    W, H = 1080, 1380
    img = Image.new('RGB', (W, H), color=(15, 15, 18))
    draw = ImageDraw.Draw(img)

    # Cores
    C_CARD = (24, 24, 29)
    C_CARD_BORDER = (45, 45, 52)
    C_TEXT = (255, 255, 255)
    C_MUTED = (161, 161, 170)
    C_AMBER = (245, 158, 11)
    C_GREEN = (16, 185, 129)
    C_BLUE = (59, 130, 246)
    C_PURPLE = (168, 85, 247)
    C_GOLD = (234, 179, 8)
    C_SLATE = (148, 163, 184)

    # Fontes do Windows
    fonte_bold = r"C:\Windows\Fonts\segoeuib.ttf" if os.path.exists(r"C:\Windows\Fonts\segoeuib.ttf") else r"C:\Windows\Fonts\arialbd.ttf"
    fonte_reg = r"C:\Windows\Fonts\segoeui.ttf" if os.path.exists(r"C:\Windows\Fonts\segoeui.ttf") else r"C:\Windows\Fonts\arial.ttf"

    f_title = ImageFont.truetype(fonte_bold, 30)
    f_badge = ImageFont.truetype(fonte_bold, 13)
    f_subtitle = ImageFont.truetype(fonte_reg, 16)
    f_kpi_label = ImageFont.truetype(fonte_reg, 15)
    f_kpi_val = ImageFont.truetype(fonte_bold, 26)
    f_kpi_sub = ImageFont.truetype(fonte_reg, 14)
    f_sec_title = ImageFont.truetype(fonte_bold, 18)
    f_item_title = ImageFont.truetype(fonte_bold, 16)
    f_item_val = ImageFont.truetype(fonte_bold, 16)
    f_item_sub = ImageFont.truetype(fonte_reg, 14)
    f_footer = ImageFont.truetype(fonte_reg, 13)

    # 1. Cabeçalho
    draw.rounded_rectangle([45, 35, 215, 63], radius=6, fill=(45, 30, 10), outline=C_AMBER, width=1)
    draw.text((58, 41), 'PAINEL EXECUTIVO', fill=C_AMBER, font=f_badge)
    draw.text((230, 41), dados.get('periodo', 'Novembro / 2026'), fill=C_MUTED, font=f_subtitle)

    draw.text((45, 75), 'PROTECBAG GRU • METAS & PROJEÇÃO', fill=C_TEXT, font=f_title)
    sub = f"{dados['dias_op']} Dias de Operação • Restam {dados['dias_rest']} dias (Novembro/2026)"
    draw.text((45, 115), sub, fill=C_MUTED, font=f_subtitle)

    draw.line([(45, 150), (1035, 150)], fill=C_CARD_BORDER, width=1)

    # 2. 4 Cards de Destaque
    cards = [
        {
            'rect': [45, 170, 520, 275],
            'title': 'FATURAMENTO REALIZADO',
            'val': dados['fat_total'],
            'sub': f"{dados['total_malas']} malas movimentadas",
            'cor_val': C_TEXT,
            'cor_bg': (18, 28, 45),
            'cor_border': (30, 58, 110)
        },
        {
            'rect': [555, 170, 1035, 275],
            'title': 'MÉDIA DIÁRIA REALIZADA',
            'val': dados['media_diaria'] + '/dia',
            'sub': 'Ritmo Médio Atual',
            'cor_val': C_GREEN,
            'cor_bg': (15, 35, 25),
            'cor_border': (20, 80, 45)
        },
        {
            'rect': [45, 290, 520, 395],
            'title': 'PROJEÇÃO FINAL DO MÊS (30 DIAS)',
            'val': dados['projecao'],
            'sub': dados['status_proj'],
            'cor_val': C_PURPLE,
            'cor_bg': (28, 18, 42),
            'cor_border': (70, 35, 110)
        },
        {
            'rect': [555, 290, 1035, 395],
            'title': 'RITMO P/ META BRONZE',
            'val': dados['media_nec_bronze'] + '/dia',
            'sub': f"Falta {dados['falta_bronze']} p/ base",
            'cor_val': C_AMBER,
            'cor_bg': (38, 24, 15),
            'cor_border': (120, 65, 20)
        }
    ]

    for c in cards:
        draw.rounded_rectangle(c['rect'], radius=12, fill=c['cor_bg'], outline=c['cor_border'], width=1)
        r = c['rect']
        draw.text((r[0] + 18, r[1] + 14), c['title'], fill=C_MUTED, font=f_kpi_label)
        draw.text((r[0] + 18, r[1] + 38), c['val'], fill=c['cor_val'], font=f_kpi_val)
        draw.text((r[0] + 18, r[1] + 74), c['sub'], fill=C_MUTED, font=f_kpi_sub)

    # 3. Metas de Faturamento
    draw.text((45, 420), '🎯 METAS DE FATURAMENTO', fill=C_TEXT, font=f_sec_title)

    metas = [
        {
            'nome': '🥉 1ª META (BRONZE) - META BASE',
            'meta': 'R$ 1.500.000,00',
            'falta': f"Falta: {dados['falta_bronze']}",
            'ritmo': f"Necessário: {dados['media_nec_bronze']}/dia",
            'pct': dados['pct_bronze'],
            'cor': C_AMBER,
            'bg': (32, 22, 15),
            'border': (80, 50, 25)
        },
        {
            'nome': '🥈 2ª META (PRATA) - INTERMEDIÁRIA',
            'meta': 'R$ 1.600.000,00',
            'falta': f"Falta: {dados['falta_prata']}",
            'ritmo': f"Necessário: {dados['media_nec_prata']}/dia",
            'pct': dados['pct_prata'],
            'cor': C_SLATE,
            'bg': (22, 26, 34),
            'border': (50, 60, 80)
        },
        {
            'nome': '🥇 3ª META (OURO) - MÁXIMA',
            'meta': 'R$ 1.650.000,00',
            'falta': f"Falta: {dados['falta_ouro']}",
            'ritmo': f"Necessário: {dados['media_nec_ouro']}/dia",
            'pct': dados['pct_ouro'],
            'cor': C_GOLD,
            'bg': (35, 30, 16),
            'border': (90, 75, 25)
        }
    ]

    y_meta = 455
    for m in metas:
        draw.rounded_rectangle([45, y_meta, 1035, y_meta + 80], radius=10, fill=m['bg'], outline=m['border'], width=1)
        draw.text((65, y_meta + 12), m['nome'], fill=m['cor'], font=f_item_title)
        draw.text((760, y_meta + 12), m['meta'], fill=C_TEXT, font=f_item_val)

        draw.text((65, y_meta + 42), m['falta'], fill=C_MUTED, font=f_item_sub)
        draw.text((380, y_meta + 42), m['ritmo'], fill=C_TEXT, font=f_item_sub)
        draw.text((760, y_meta + 42), f"{m['pct']} alcançado", fill=m['cor'], font=f_item_sub)
        y_meta += 92

    # 4. Desempenho por Turno
    draw.text((45, y_meta + 15), '👥 PARTICIPAÇÃO POR TURNO', fill=C_TEXT, font=f_sec_title)
    y_t = y_meta + 50

    turnos = [
        {
            'nome': '☀️ TURNO B (Tarde) - DESTAQUE 🔥',
            'fat': dados['fat_b'],
            'malas': f"{dados['malas_b']} malas",
            'part': f"{dados['pct_b']}% do faturamento",
            'media': f"Média: {dados['media_b']}/dia",
            'cor': C_AMBER,
            'bg': (32, 24, 15),
            'border': (90, 60, 20)
        },
        {
            'nome': '🌅 TURNO A (Manhã)',
            'fat': dados['fat_a'],
            'malas': f"{dados['malas_a']} malas",
            'part': f"{dados['pct_a']}% do faturamento",
            'media': f"Média: {dados['media_a']}/dia",
            'cor': C_BLUE,
            'bg': (18, 25, 38),
            'border': (40, 60, 95)
        },
        {
            'nome': '🌙 TURNO C (Madrugada)',
            'fat': dados['fat_d'],
            'malas': f"{dados['malas_d']} malas",
            'part': f"{dados['pct_d']}% do faturamento",
            'media': f"Média: {dados['media_d']}/dia",
            'cor': C_PURPLE,
            'bg': (28, 18, 38),
            'border': (65, 40, 90)
        }
    ]

    for t in turnos:
        draw.rounded_rectangle([45, y_t, 1035, y_t + 80], radius=10, fill=t['bg'], outline=t['border'], width=1)
        draw.text((65, y_t + 12), t['nome'], fill=t['cor'], font=f_item_title)
        draw.text((760, y_t + 12), t['fat'], fill=C_TEXT, font=f_item_val)

        draw.text((65, y_t + 42), t['part'], fill=C_TEXT, font=f_item_sub)
        draw.text((380, y_t + 42), t['media'], fill=C_MUTED, font=f_item_sub)
        draw.text((760, y_t + 42), t['malas'], fill=C_MUTED, font=f_item_sub)
        y_t += 92

    # 5. Rodapé
    draw.line([(45, y_t + 15), (1035, y_t + 15)], fill=C_CARD_BORDER, width=1)
    foot = f"Total Geral: {dados['total_malas']} malas • Protecbag GRU Aeroporto • Gerado em {dados.get('data_hora', '')}"
    draw.text((45, y_t + 30), foot, fill=C_MUTED, font=f_footer)

    img.save(caminho_saida, quality=95)
    print(f"🖼️ Imagem gerada: {caminho_saida}")

    # Copiar imagem para a área de transferência do Windows via PowerShell
    try:
        cmd = f"powershell -NoProfile -Command \"Add-Type -AssemblyName System.Windows.Forms; [System.Windows.Forms.Clipboard]::SetImage([System.Drawing.Image]::FromFile('{caminho_saida}'))\""
        subprocess.run(cmd, shell=True, check=False)
        print("📋 Imagem copiada para a Área de Transferência (Pronta para Ctrl+V no WhatsApp)!")
    except Exception as e:
        print(f"⚠️ Aviso ao copiar imagem: {e}")

if __name__ == '__main__':
    # Teste rápido
    dados_demo = {
        'periodo': 'Novembro / 2026',
        'dias_op': 13,
        'dias_rest': 17,
        'fat_total': 'R$ 599.651,00',
        'total_malas': '3.508',
        'media_diaria': 'R$ 46.127,00',
        'projecao': 'R$ 1.383.810,00',
        'status_proj': 'Abaixo da Meta Bronze (-R$ 116.190,00)',
        'falta_bronze': 'R$ 900.349,00',
        'media_nec_bronze': 'R$ 52.961,71',
        'pct_bronze': '39,98%',
        'falta_prata': 'R$ 1.000.349,00',
        'media_nec_prata': 'R$ 58.844,06',
        'pct_prata': '37,48%',
        'falta_ouro': 'R$ 1.050.349,00',
        'media_nec_ouro': 'R$ 61.785,24',
        'pct_ouro': '36,34%',
        'fat_b': 'R$ 316.307,00',
        'malas_b': '1.845',
        'pct_b': '52,75',
        'media_b': 'R$ 24.331,31',
        'fat_a': 'R$ 229.384,00',
        'malas_a': '1.337',
        'pct_a': '38,25',
        'media_a': 'R$ 17.644,92',
        'fat_d': 'R$ 53.960,00',
        'malas_d': '326',
        'pct_d': '9,00',
        'media_d': 'R$ 4.150,77',
        'data_hora': '13/11/2026 14:15'
    }
    saida = os.path.join(os.path.expanduser('~'), 'Desktop', 'Relatorio_Diretoria_GRU.png')
    gerar_card_imagem(dados_demo, saida)
