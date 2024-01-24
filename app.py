# A biblioteca openpyxl foi utilizada para a leitura da planilha
import openpyxl
# A biblioteca Pillow foi utilizada para abrir e desenhar na imagem
from PIL import Image, ImageDraw, ImageFont
# A biblioteca sleep foi utilizada só para garantir que não ocorra nenhum erro de sobreposição do texto pela velocidade de execução do código
from time import sleep

# Lendo a planilha
planilha = openpyxl.load_workbook('planilha_com_dados/planilha_dos_alunos.xlsx')
sheet_alunos = planilha['Sheet1']

# Acessando as linhas e desenhando no certificado
for linha in sheet_alunos.iter_rows(min_row=2):
    nome_curso = linha[0].value
    nome_participante = linha[1].value
    data_inicio = str(linha[2].value)
    data_termino = str(linha[3].value)
    carga_horaria = str(linha[4].value) + 'h'
    data_emissao = str(linha[5].value)

    # Importando fontes
    fonte_nome = ImageFont.truetype('fontes_usadas/tahomabd.ttf', 125)
    fonte_geral = ImageFont.truetype('fontes_usadas/tahoma.ttf', 110)

    # Importando imagem (certificado)
    certificado = Image.open('imagem_original_certificado/certificado_para_usar.png')
    desenhar_certificado = ImageDraw.Draw(certificado)

    desenhar_certificado.text((930, 900), nome_participante, fill='black', font=fonte_nome) #nome participante
    desenhar_certificado.text((975, 1135), nome_curso, fill='black', font=fonte_geral) #nome curso
    desenhar_certificado.text((1450, 1350), carga_horaria, fill='black', font=fonte_geral) #carga horaria
    desenhar_certificado.text((900, 1805), data_inicio, fill='black', font=fonte_geral) #data inicio
    desenhar_certificado.text((900, 2095), data_termino, fill='black', font=fonte_geral) #data termino
    desenhar_certificado.text((2250, 2095), data_emissao, fill='black', font=fonte_geral) #data emissão
    
    certificado.save(f'certificados_salvos/{nome_participante}.png')

    sleep(1)
