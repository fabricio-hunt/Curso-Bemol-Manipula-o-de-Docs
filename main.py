import pandas as pd
from docxtpl import DocxTemplate
from docx2pdf import convert
from datetime import datetime
import os
 
# Leitura do Excel --> dataframe
df = pd.read_excel("exemplo_dados.xlsx")
 
# Tradutor simples para meses em português
meses_pt = {
    'January': 'janeiro',
    'February': 'fevereiro',
    'March': 'março',
    'April': 'abril',
    'May': 'maio',
    'June': 'junho',
    'July': 'julho',
    'August': 'agosto',
    'September': 'setembro',
    'October': 'outubro',
    'November': 'novembro',
    'December': 'dezembro'
}
 
# Iterando linha a linha
for i, row in df.iterrows():
    # Pegando as variáveis que serão substituídas
    nome = row['NOME']
    cpf = row['CPF']
    rg = row['RG']
    endereco = row['ENDEREÇO']
 
    # Separando endereço: "rua, numero - bairro"
    parte1, bairro = endereco.split(' - ')
    rua, numero = parte1.split(', ')
 
    # Limpando espaços
    rua = rua.strip()
    numero = numero.strip()
    bairro = bairro.strip()
 
    # Pegando a data atual
    data_hoje = datetime.today()
    dia = data_hoje.day
    mes_en = data_hoje.strftime('%B')
    ano = data_hoje.year
    mes_pt = meses_pt[mes_en] # meses_pt['June] --> junho
 
    # Preenchendo o modelo
    doc = DocxTemplate("MODELO-DECLARAÇÃO-DE-DISPONIBILIDADE-DE-HORÁRIO.docx")
    context = {
        "NOME": nome,
        "NUMEROCPF": cpf,
        "NUMERORG": rg,
        "NOMERUA": rua,
        "NUMEROCASA": numero,
        "NOMEBAIRRO": bairro,
        "DIA": dia,
        "MÊS": mes_pt,
        "ANO": ano
    }
    doc.render(context)
 
    # Salvando temporariamente
    temp_docx = f"declaracao_{nome}.docx"
    temp_pdf = f"declaracao_{nome}.pdf"
    doc.save(temp_docx)
 
    # Convertendo para PDF
    convert(temp_docx, temp_pdf)
 
    # (Opcional) remover docx
    os.remove(temp_docx)
 
    print(f"✅ PDF gerado: {temp_pdf}")