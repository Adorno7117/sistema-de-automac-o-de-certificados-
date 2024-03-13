import openpyxl 
from PIL import Image, ImageDraw, ImageFont

workbookAlunos = openpyxl.load_workbook('planilha_alunos.xlsx')
sheetAlunos = workbookAlunos['Sheet1']

# Extrair dados

for indice, linha in enumerate(sheetAlunos.iter_rows(min_row=2)):
    nomeCurso = linha[0].value
    nomeParticipante = linha[1].value
    tipoParticipante = linha[2].value
    dataInicio = linha[3].value
    dataFinal = linha[4].value
    cargaHoraria = linha[5].value
    dataEmissao = linha[6].value

    fonteNome = ImageFont.truetype('./tahomabd.ttf',90)
    fonteNormal = ImageFont.truetype('./tahoma.ttf',80)
    fonteData = ImageFont.truetype('./tahoma.ttf',55)

    image = Image.open('./certificado_padrao.jpg')
    escrever = ImageDraw.Draw(image)

    escrever.text((1020,827), nomeParticipante, fill='black', font=fonteNome)
    escrever.text((1060,950), nomeCurso, fill='black', font=fonteNormal)
    escrever.text((1435,1065), tipoParticipante, fill='black', font=fonteNormal)
    escrever.text((1480,1182), str(cargaHoraria), fill='black', font=fonteNormal)
    
    escrever.text((750,1770), dataInicio, fill='black', font=fonteData)
    escrever.text((750,1930), dataFinal, fill='black', font=fonteData)
    escrever.text((2220,1930), dataEmissao, fill='black', font=fonteData)
    

    image.save(f'./Certificados/{indice}_{nomeParticipante}_certificado.png')