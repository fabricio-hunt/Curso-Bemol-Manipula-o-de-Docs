# 📝 Gerador de Declarações Personalizadas em PDF

Este projeto automatiza a geração de declarações personalizadas a partir de um modelo `.docx`, com dados provenientes de uma planilha Excel. Cada documento é preenchido com informações como nome, CPF, RG e endereço, e convertido automaticamente para o formato PDF.

---

## 📌 Funcionalidades

- Leitura de dados de um arquivo Excel (`.xlsx`)
- Preenchimento automático de um modelo Word com os dados individuais
- Conversão dos arquivos `.docx` gerados para o formato `.pdf`
- Tradução automática do mês atual para português
- Geração de arquivos PDF nomeados com base no nome da pessoa

---

## 📁 Estrutura Esperada

📂 seu-projeto/
├── gerar_declaracoes.py
├── exemplo_dados.xlsx
├── MODELO-DECLARAÇÃO-DE-DISPONIBILIDADE-DE-HORÁRIO.docx


---

## 📊 Exemplo de Dados (Excel)

A planilha `exemplo_dados.xlsx` deve conter a seguinte estrutura:

| NOME          | CPF           | RG            | ENDEREÇO                      |
|---------------|---------------|---------------|-------------------------------|
| João da Silva | 000.000.000-00| 123456789      | Rua Exemplo, 123 - Centro     |

---

## 🧾 Modelo `.docx`

O modelo Word deve conter os seguintes campos com chaves duplas, compatíveis com o template do `docxtpl`:

- `{{NOME}}`
- `{{NUMEROCPF}}`
- `{{NUMERORG}}`
- `{{NOMERUA}}`
- `{{NUMEROCASA}}`
- `{{NOMEBAIRRO}}`
- `{{DIA}}`
- `{{MÊS}}`
- `{{ANO}}`

---

## 🛠️ Requisitos

- Python 3.7 ou superior
- Sistema operacional Windows (necessário para `docx2pdf`)
- Pacotes Python:

```bash
pip install pandas docxtpl docx2pdf
