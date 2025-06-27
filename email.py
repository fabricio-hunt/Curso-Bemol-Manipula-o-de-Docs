import os
import pandas as pd
import win32com.client as win32

# CONFIGURAÇÃO
EXCEL_PATH = "exemplo_dados.xlsx"
PDF_DIR    = "."  # ou "declaracoes_output" se você usar pasta específica
MEU_EMAIL  = "fabriciomacedo@bemol.com.br"

def main():
    # carrega a lista de nomes + e-mails
    df = pd.read_excel(EXCEL_PATH)
    outlook = win32.Dispatch('Outlook.Application')

    for _, row in df.iterrows():
        nome       = row['NOME']
        email_dest = row['E-MAIL']
        # monta o nome do pdf exatamente como está na pasta
        pdf_name   = f"declaracao_{nome}.pdf"
        pdf_path   = os.path.join(PDF_DIR, pdf_name)

        if not os.path.isfile(pdf_path):
            print(f"⚠️  {pdf_name} não encontrado, pulando {nome}")
            continue

        # monta e envia o e-mail
        mail = outlook.CreateItem(0)
        mail.Subject  = "[Declaração RH] - Disponibilidade de Horário"
        mail.HTMLBody = f"""
            <p>Olá {nome},</p>
            <p>Conforme solicitado, segue em anexo sua declaração de disponibilidade de horário.</p>
            <p>Atenciosamente,<br>Equipe de RH</p>
        """
        mail.To = email_dest
        mail.CC = MEU_EMAIL
        mail.Attachments.Add(os.path.abspath(pdf_path))
        mail.Send()

        print(f"✅ Enviado para {nome} <{email_dest}>")

    print("\n→ Envio concluído.")

if __name__ == "__main__":
    main()
