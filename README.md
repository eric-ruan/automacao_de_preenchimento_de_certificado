# Estudo e Aprimoramento em Python

### O objetivo do código é ler uma planilha com os seguintes dados: nome do curso, nome do participante, data de início, data de término, carga horária(horas) e data de emissão. Após a leitura, o objetivo final é gravar esses dados em um certificado, isso de forma automática para agilizar o trabalho, já que a planilha possui 100 dados. Para praticar meus conhecimentos em edição de imagem o certificado foi criado por mim(sem nenhuma validade).

#### Este código serve como uma forma de estudo e aprimoramento das habilidades em Python. Os dados utilizados na planilha são fictícios, assim como o certificado, que foi criado com o propósito de ter uma imagem base para utilizar no código. O certificado não possui nenhuma validade real.

### A biblioteca openpyxl foi utilizada para a leitura da planilha
```python
import openpyxl
```
### A biblioteca Pillow foi utilizada para abrir e desenhar na imagem
```python
from PIL import Image, ImageDraw, ImageFont
```
### A biblioteca sleep foi utilizada para garantir que não ocorra nenhum erro de sobreposição do texto pela velocidade de execução do código
```python
from time import sleep
```

### Lendo a planilha
```python
planilha = openpyxl.load_workbook('planilha_com_dados/planilha_dos_alunos.xlsx')
sheet_alunos = planilha['Sheet1']
```

### Acessando as linhas e desenhando no certificado
```python
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

    # Desenhando no certificado
    desenhar_certificado.text((930, 900), nome_participante, fill='black', font=fonte_nome)  # Nome do participante
    desenhar_certificado.text((975, 1135), nome_curso, fill='black', font=fonte_geral)  # Nome do curso
    desenhar_certificado.text((1450, 1350), carga_horaria, fill='black', font=fonte_geral)  # Carga horária
    desenhar_certificado.text((900, 1805), data_inicio, fill='black', font=fonte_geral)  # Data de início
    desenhar_certificado.text((900, 2095), data_termino, fill='black', font=fonte_geral)  # Data de término
    desenhar_certificado.text((2250, 2095), data_emissao, fill='black', font=fonte_geral)  # Data de emissão

    certificado.save(f'certificados_salvos/{nome_participante}.png')

    sleep(1)
```
