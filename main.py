import pandas as pd
from datetime import datetime
import os
from sendgrid import SendGridAPIClient
from sendgrid.helpers.mail import Mail

# ===== CONFIGURAÇÕES =====
API_KEY = os.getenv("SENDGRID_API_KEY")
EMAIL_REMETENTE = "daniel.cardoso@potenza-investimentos.com"

CAMINHO_ARQUIVO = "aniversarios.xlsx"

# ===== LER PLANILHA =====
df = pd.read_excel(CAMINHO_ARQUIVO, sheet_name=0)

hoje = datetime.now().strftime("%d/%m")

# ===== FUNÇÃO PARA MULTI-EMPRESA =====
def pertence_empresa(empresas_str, empresa_alvo):
    if pd.isna(empresas_str):
        return False
    lista = [e.strip() for e in str(empresas_str).split(",")]
    return empresa_alvo in lista

# ===== FUNÇÃO DE ENVIO COM BCC =====
def enviar_email_com_bcc(destino, bcc_lista, assunto, corpo):

    if not destino:
        return

    message = Mail(
        from_email=EMAIL_REMETENTE,
        to_emails=destino,
        subject=assunto,
        html_content=corpo
    )

    if bcc_lista:
        for email in bcc_lista:
            if email and email != destino:
                message.add_bcc(email)

    try:
        sg = SendGridAPIClient(API_KEY)
        sg.send(message)
        print(f"Email enviado para {destino}")
    except Exception as e:
        print(f"Erro ao enviar para {destino}: {e}")

# ===== PROCESSAR DADOS =====
for _, row in df.iterrows():

    if row.isnull().all():
        continue

    nome = row.get('Nome')
    email = row.get('Email')
    empresa = row.get('Empresa')
    tipo = row.get('Tipo')

    if pd.isna(nome) or pd.isna(email) or pd.isna(row.get('DataNascimento')):
        continue

    try:
        data = pd.to_datetime(row.get('DataNascimento')).strftime("%d/%m")
    except:
        continue

    if data == hoje and tipo == "Pessoa":

        # 👥 filtra equipe considerando múltiplas empresas
        equipe = df[df['Empresa'].apply(lambda x: pertence_empresa(x, empresa))]

        # lista de emails da equipe (remove aniversariante)
        emails_equipe = [
            e for e in equipe['Email'].dropna().tolist()
            if e != email
        ]

        # 🎉 envia email
        enviar_email_com_bcc(
            destino=email,
            bcc_lista=emails_equipe,
            assunto=f"🎉 Feliz Aniversário, {nome}!",
            corpo=f"""
<div style="font-family: Arial, sans-serif; color: #333; line-height: 1.6;">

<p>Olá <b>{nome}</b>,</p>

<p>Hoje é um dia especial — e nós não poderíamos deixar de celebrar com você! 🎉</p>

<p>
Toda a equipe da <b>{empresa}</b> deseja um aniversário repleto de alegria, saúde e conquistas.
</p>

<p>
Que este novo ciclo venha com ainda mais sucesso, realizações e bons momentos — dentro e fora do trabalho.
</p>

<p>
Agradecemos por fazer parte da nossa equipe e por tudo que você constrói conosco todos os dias.
</p>

<p>🎂 <b>Feliz aniversário!</b></p>

<br>

<p>Com carinho,<br>
<b>Equipe {empresa}</b></p>

</div>
"""
        )

print("Processo finalizado com sucesso.")
