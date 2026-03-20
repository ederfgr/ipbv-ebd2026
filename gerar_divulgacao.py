#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from PIL import Image, ImageDraw, ImageFont
from textwrap import wrap
import os

# Cores Pedras Vivas (conforme manual de identidade)
COR_AZUL = (44, 55, 104)  # #2C3768
COR_MARROM = (103, 89, 75)  # #67594B
COR_LARANJA = (240, 127, 37)  # #F07F25
COR_DOURADO = (209, 184, 74)  # #D1B84A
COR_BRANCO = (255, 255, 255)
COR_CINZA_CLARO = (240, 240, 240)
COR_CINZA_ESCURO = (80, 80, 80)

# Cores para as aulas (paleta diversa)
CORES_AULAS = [
    (52, 152, 219),   # Azul
    (231, 76, 60),    # Vermelho
    (39, 174, 96),    # Verde
    (241, 196, 15),   # Amarelo
    (155, 89, 182),   # Roxo
    (26, 188, 156),   # Turquesa
    (230, 126, 34),   # Laranja
    (52, 73, 94),     # Azul escuro
    (189, 195, 199),  # Cinza
    (155, 89, 182),   # Roxo 2
    (230, 126, 34),   # Laranja 2
    (46, 204, 113),   # Verde 2
    (52, 152, 219),   # Azul 2
    (231, 76, 60),    # Vermelho 2
    (243, 156, 18),   # Ouro
    (149, 165, 166),  # Cinza 2
    (44, 62, 80),     # Azul escuro 2
    (26, 188, 156),   # Turquesa 2
]

# Dados das aulas
AULAS_TOTAL = [
    ("NOSSA HISTÓRIA E IDENTIDADE", [
        "Aula 1: Nossa História",
        "Aula 2: Como Nos Organizamos",
        "Aula 3: Como Crescemos Espiritualmente"
    ]),
    ("INTRODUÇÃO BÍBLICA", [
        "Aula 4: A Palavra de Deus - Nosso Fundamento"
    ]),
    ("A DOUTRINA DE DEUS", [
        "Aula 5: A Trindade - Um Só Deus em Três Pessoas",
        "Aula 6: Os Atributos de Deus"
    ]),
    ("A DOUTRINA DO HOMEM", [
        "Aula 7: A Criação do Homem - Feitos à Imagem de Deus",
        "Aula 8: Livre-Arbítrio e a Queda do Homem"
    ]),
    ("A DOUTRINA DO PECADO", [
        "Aula 9: A Doutrina do Pecado"
    ]),
    ("A DOUTRINA DE JESUS CRISTO", [
        "Aula 10: Jesus Cristo - Plenamente Deus e Plenamente Homem",
        "Aula 11: Jesus Cristo - Humilhação, Exaltação e Seus Ofícios"
    ]),
    ("A DOUTRINA DO ESPÍRITO SANTO", [
        "Aula 12: O Espírito Santo - Pessoa e Obra",
        "Aula 13: Os Dons e o Fruto do Espírito Santo"
    ]),
    ("A DOUTRINA DA SALVAÇÃO", [
        "Aula 14: A Doutrina da Salvação - Parte 1",
        "Aula 15: A Doutrina da Salvação - Parte 2"
    ]),
    ("A DOUTRINA DAS ÚLTIMAS COISAS", [
        "Aula 16: A Doutrina das Últimas Coisas"
    ]),
    ("A DOUTRINA DA IGREJA", [
        "Aula 17: A Doutrina da Igreja - Parte 1",
        "Aula 18: A Doutrina da Igreja - Parte 2"
    ])
]

def desenhar_formas_decorativas(draw, x, y, tamanho, cor):
    """Desenha formas decorativas geométricas"""
    # Triângulo estilizado
    pontos = [
        (x, y - tamanho),
        (x + tamanho, y + tamanho // 2),
        (x - tamanho, y + tamanho // 2)
    ]
    draw.polygon(pontos, fill=cor)

def criar_background_gradiente(img, cor1, cor2):
    """Cria um efeito de gradiente no background"""
    pixels = img.load()
    largura, altura = img.size
    
    for y in range(altura):
        r = int(cor1[0] + (cor2[0] - cor1[0]) * y / altura)
        g = int(cor1[1] + (cor2[1] - cor1[1]) * y / altura)
        b = int(cor1[2] + (cor2[2] - cor1[2]) * y / altura)
        
        for x in range(largura):
            pixels[x, y] = (r, g, b)

def criar_material_divulgacao():
    """Cria o material de divulgação em A4 (2100x2970 pixels a 300 DPI)"""
    
    # A4 em pixels a 300 DPI: 2100x2970
    largura = 2100
    altura = 2970
    
    # Criar imagem com fundo branco
    img = Image.new('RGB', (largura, altura), COR_BRANCO)
    draw = ImageDraw.Draw(img)
    
    # Tentar carregar fontes
    try:
        fonte_titulo_grande = ImageFont.truetype("arial.ttf", 140)
        fonte_titulo = ImageFont.truetype("arial.ttf", 90)
        fonte_subtitulo = ImageFont.truetype("arial.ttf", 55)
        fonte_secao = ImageFont.truetype("arial.ttf", 45)
        fonte_aula = ImageFont.truetype("arial.ttf", 42)  # Aumentada
        fonte_pequeno = ImageFont.truetype("arial.ttf", 35)
        fonte_versaculo = ImageFont.truetype("arial.ttf", 36)
    except:
        fonte_titulo_grande = ImageFont.load_default()
        fonte_titulo = ImageFont.load_default()
        fonte_subtitulo = ImageFont.load_default()
        fonte_secao = ImageFont.load_default()
        fonte_aula = ImageFont.load_default()
        fonte_pequeno = ImageFont.load_default()
        fonte_versaculo = ImageFont.load_default()
    
    # Faixa colorida superior (gradiente simulado)
    for i in range(0, 240, 5):
        cor_gradiente = (
            int(COR_AZUL[0] - (COR_AZUL[0] - COR_MARROM[0]) * i / 240),
            int(COR_AZUL[1] - (COR_AZUL[1] - COR_MARROM[1]) * i / 240),
            int(COR_AZUL[2] - (COR_AZUL[2] - COR_MARROM[2]) * i / 240)
        )
        draw.rectangle([(0, i), (largura, i + 5)], fill=cor_gradiente)
    
    # Triângulos decorativos no topo
    desenhar_formas_decorativas(draw, 120, 80, 35, COR_DOURADO)
    desenhar_formas_decorativas(draw, largura - 120, 80, 35, COR_DOURADO)
    
    # Título principal
    titulo = "O QUE CREMOS"
    titulo_bbox = draw.textbbox((0, 0), titulo, font=fonte_titulo_grande)
    titulo_width = titulo_bbox[2] - titulo_bbox[0]
    titulo_x = (largura - titulo_width) // 2
    draw.text((titulo_x, 45), titulo, fill=COR_DOURADO, font=fonte_titulo_grande)
    
    y_pos = 260
    
    # Subtítulo
    subtitulo = "Fundamentos Doutrinários da Fé Cristã"
    subtitulo_bbox = draw.textbbox((0, 0), subtitulo, font=fonte_subtitulo)
    subtitulo_width = subtitulo_bbox[2] - subtitulo_bbox[0]
    subtitulo_x = (largura - subtitulo_width) // 2
    draw.text((subtitulo_x, y_pos), subtitulo, fill=COR_AZUL, font=fonte_subtitulo)
    
    y_pos += 100
    
    # Data de início com caixa decorativa
    data = "Início: 01/03/2026"
    draw.rectangle([(largura//2 - 300, y_pos - 15), (largura//2 + 300, y_pos + 55)], 
                   fill=COR_LARANJA, outline=COR_AZUL, width=4)
    data_bbox = draw.textbbox((0, 0), data, font=fonte_pequeno)
    data_width = data_bbox[2] - data_bbox[0]
    data_x = (largura - data_width) // 2
    draw.text((data_x, y_pos), data, fill=COR_BRANCO, font=fonte_pequeno)
    
    y_pos += 120
    
    # Linha decorativa em dois segmentos
    draw.rectangle([(200, y_pos), (largura//2 - 50, y_pos + 6)], fill=COR_LARANJA)
    draw.rectangle([(largura//2 + 50, y_pos), (largura - 200, y_pos + 6)], fill=COR_LARANJA)
    
    # Texto central
    circulo_y = y_pos + 20
    draw.ellipse([(largura//2 - 25, circulo_y - 25), (largura//2 + 25, circulo_y + 25)], 
                 fill=COR_DOURADO)
    
    y_pos += 80
    
    # Separar em duas colunas
    margem_lateral = 80
    espaco_coluna = largura // 2
    col1_x = margem_lateral
    col2_x = largura // 2 + 40
    
    # Dividir aulas em duas listas
    aulas_col1 = []
    aulas_col2 = []
    contador_total = 0
    
    for secao, aulas in AULAS_TOTAL:
        for aula in aulas:
            if contador_total < 9:
                aulas_col1.append((secao if contador_total == sum(len(a[1]) for a in AULAS_TOTAL[:AULAS_TOTAL.index([s for s in AULAS_TOTAL if s[0] == secao][0])]) else None, aula))
            else:
                aulas_col2.append((secao if contador_total == sum(len(a[1]) for a in AULAS_TOTAL[:AULAS_TOTAL.index([s for s in AULAS_TOTAL if s[0] == secao][0])]) else None, aula))
            contador_total += 1
    
    # Simplificar: pegar todas as aulas e dividir direto
    todas_aulas = []
    for secao, aulas in AULAS_TOTAL:
        for aula in aulas:
            todas_aulas.append(aula)
    
    meio = len(todas_aulas) // 2
    aulas_col1 = todas_aulas[:meio]
    aulas_col2 = todas_aulas[meio:]
    
    # Desenhar coluna 1
    y_col1 = y_pos + 40
    for idx, aula in enumerate(aulas_col1):
        cor_aula = CORES_AULAS[idx % len(CORES_AULAS)]
        
        # Linha colorida à esquerda
        draw.rectangle([(col1_x - 20, y_col1), (col1_x - 8, y_col1 + 65)], fill=cor_aula)
        
        # Número da aula com fundo colorido
        num_match = aula.split(':')[0].strip()
        draw.rectangle([(col1_x, y_col1 - 5), (col1_x + 70, y_col1 + 60)], 
                       fill=cor_aula, outline=COR_BRANCO, width=2)
        draw.text((col1_x + 15, y_col1 + 5), num_match, fill=COR_BRANCO, font=fonte_secao)
        
        # Texto da aula
        titulo_aula = ': '.join(aula.split(':')[1:]).strip()
        texto_wrapped = wrap(titulo_aula, width=35)
        y_linha = y_col1 + 5
        for linha in texto_wrapped:
            draw.text((col1_x + 90, y_linha), linha, fill=COR_CINZA_ESCURO, font=fonte_aula)
            y_linha += 55
        
        y_col1 = max(y_linha + 15, y_col1 + 75)
    
    # Desenhar coluna 2
    y_col2 = y_pos + 40
    for idx, aula in enumerate(aulas_col2):
        cor_aula = CORES_AULAS[(idx + meio) % len(CORES_AULAS)]
        
        # Linha colorida à esquerda
        draw.rectangle([(col2_x - 20, y_col2), (col2_x - 8, y_col2 + 65)], fill=cor_aula)
        
        # Número da aula com fundo colorido
        num_match = aula.split(':')[0].strip()
        draw.rectangle([(col2_x, y_col2 - 5), (col2_x + 70, y_col2 + 60)], 
                       fill=cor_aula, outline=COR_BRANCO, width=2)
        draw.text((col2_x + 15, y_col2 + 5), num_match, fill=COR_BRANCO, font=fonte_secao)
        
        # Texto da aula
        titulo_aula = ': '.join(aula.split(':')[1:]).strip()
        texto_wrapped = wrap(titulo_aula, width=35)
        y_linha = y_col2 + 5
        for linha in texto_wrapped:
            draw.text((col2_x + 90, y_linha), titulo_aula if len(texto_wrapped) == 1 else linha, 
                     fill=COR_CINZA_ESCURO, font=fonte_aula)
            y_linha += 55
            if len(texto_wrapped) > 1:
                titulo_aula = ""
        
        y_col2 = max(y_linha + 15, y_col2 + 75)
    
    # Posição final
    y_final = max(y_col1, y_col2) + 60
    
    # Linha decorativa
    draw.rectangle([(200, y_final), (largura - 200, y_final + 4)], fill=COR_MARROM)
    
    y_final += 50
    
    # Versículo no rodapé
    versaculo = '"Vocês, também, como pedras vivas, deixem que Deus os use'
    versaculo2 = 'na construção de um templo espiritual..." - 1 Pedro 2:5'
    
    verso1_bbox = draw.textbbox((0, 0), versaculo, font=fonte_versaculo)
    verso1_width = verso1_bbox[2] - verso1_bbox[0]
    verso1_x = (largura - verso1_width) // 2
    draw.text((verso1_x, y_final), versaculo, fill=COR_AZUL, font=fonte_versaculo)
    
    y_final += 55
    
    verso2_bbox = draw.textbbox((0, 0), versaculo2, font=fonte_versaculo)
    verso2_width = verso2_bbox[2] - verso2_bbox[0]
    verso2_x = (largura - verso2_width) // 2
    draw.text((verso2_x, y_final), versaculo2, fill=COR_AZUL, font=fonte_versaculo)
    
    # Faixa colorida inferior (gradiente simulado)
    y_footer = altura - 180
    for i in range(0, 180, 5):
        cor_gradiente = (
            int(COR_AZUL[0] - (COR_AZUL[0] - COR_MARROM[0]) * i / 180),
            int(COR_AZUL[1] - (COR_AZUL[1] - COR_MARROM[1]) * i / 180),
            int(COR_AZUL[2] - (COR_AZUL[2] - COR_MARROM[2]) * i / 180)
        )
        draw.rectangle([(0, y_footer + i), (largura, y_footer + i + 5)], fill=cor_gradiente)
    
    # Triângulos decorativos no rodapé
    desenhar_formas_decorativas(draw, 120, altura - 90, 35, COR_DOURADO)
    desenhar_formas_decorativas(draw, largura - 120, altura - 90, 35, COR_DOURADO)
    
    # Salvar imagem
    output_path = r'c:\Users\ederf\OneDrive\Eder\ibpv\EBD-2026\docs\EBD-2026-Material-Divulgacao.png'
    img.save(output_path, 'PNG', quality=95, dpi=(300, 300))
    
    print(f"[OK] Material de divulgacao criado: {output_path}")
    print(f"Tamanho: A4 (2100x2970 px a 300 DPI)")
    print(f"Layout: 2 Colunas com fonte aumentada")
    print(f"Elementos: Gradientes, formas geométricas, caixas coloridas por aula")

if __name__ == '__main__':
    criar_material_divulgacao()
