import pandas as pd
import requests
import os
from datetime import datetime, timedelta
import unicodedata


# =========================
# CONFIGURAÇÕES
# =========================
URL_PLANILHA = (
    "https://docs.google.com/spreadsheets/d/"
    "1A0beFGh1PL-t7PTuZvRRuuk-nDQeWZxsMPVQ1I4QM0I"
    "/gviz/tq?tqx=out:csv&sheet=Pedidos"
)

DIAS_ALERTA = 7
BREVO_API_KEY = os.getenv("BREVO_API_KEY")

EMAIL_REMETENTE = {
    "name": "Quimlab",
    "email": "quimlabcomercial@gmail.com"  # validado na Brevo
}

EMAIL_DESTINATARIO = {
    "email": "marcos@quimlab.com.br",
    "name": "Marcos"
}


# =========================
# FUNÇÕES AUXILIARES
# =========================
def normalizar_texto(texto):
    if pd.isna(texto):
        return ""
    texto = str(texto).lower().strip()
    texto = unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode("ascii")
    return texto


def obter_data_final(valor):
    if pd.isna(valor):
        return None

    texto = str(valor)

    if "à" in texto:
        _, fim = texto.split("à")
        texto = fim.strip()

    try:
        return pd.to_datetime(texto, dayfirst=True).date()
    except Exception:
        return None


def dias_uteis_entre(hoje, data_final):
    dias = 0
    data = hoje

    while data < data_final:
        data += timedelta(days=1)
        if data.weekday() < 5:  # segunda a sexta
            dias += 1

    return dias



# def interpretar_data(valor):
#     if pd.isna(valor):
#         return None

#     texto = str(valor)

#     # Ex: 10/12/2025 à 16/12/2025
#     if "à" in texto:
#         texto = texto.split("à")[-1].strip()

#     try:
#         return pd.to_datetime(texto, dayfirst=True).date()
#     except Exception:
#         return None


# =========================
# ENVIO DE EMAIL (BREVO)
# =========================
def enviar_email_brevo(resultados):
    print("📨 Enviando e-mail via Brevo")

    linhas_html = ""
    for r in resultados:
        linhas_html += f"""
        <tr>
            <td>{r['data']}</td>
            <td>{r['of']}</td>
            <td>{r['status']}</td>
            <td>{r['cliente']}</td>
            <td>{r['setor']}</td>
        </tr>
        """

    html = f"""
    <html>
    <body>
        <p>Ordens de Fabricação com atenção:</p>
        <table border="1" cellpadding="5" cellspacing="0">
            <tr>
                <th>Data</th>
                <th>OF</th>
                <th>Status</th>
                <th>Cliente</th>
                <th>Setor</th>
            </tr>
            {linhas_html}
        </table>
    </body>
    </html>
    """

    payload = {
        "sender": EMAIL_REMETENTE,
        "to": [EMAIL_DESTINATARIO],
        "subject": "⚠️ Alerta de Ordens de Fabricação",
        "htmlContent": html
    }

    headers = {
        "api-key": BREVO_API_KEY,
        "content-type": "application/json",
        "accept": "application/json"
    }

    response = requests.post(
        "https://api.brevo.com/v3/smtp/email",
        json=payload,
        headers=headers
    )

    print("📧 Status Brevo:", response.status_code)
    print("📧 Resposta Brevo:", response.text)


# =========================
# FUNÇÃO PRINCIPAL
# =========================
def rodar_verificacao():
    print("🚀 Script iniciado")

    hoje = datetime.today().date()
    limite = hoje + timedelta(days=DIAS_ALERTA)

    print("🌐 Lendo planilha Google Sheets (CSV)")
    df = pd.read_csv(URL_PLANILHA)

    print("📥 Linhas carregadas:", len(df))

    resultados = []

hoje = datetime.today().date()

for _, linha in df.iterrows():

    data_final = obter_data_final(linha["Data"])
    if not data_final:
        continue

    status_original = str(linha["Status"])
    status = normalizar_texto(status_original)

    # Apenas "Em Produção"
    if "produc" not in status:
        continue

    dias_uteis = dias_uteis_entre(hoje, data_final)

    print(
        f"OF {linha['OF']} | Data final: {data_final} | Dias úteis restantes: {dias_uteis}"
    )

    # ✅ REGRA CORRETA
    if dias_uteis <= DIAS_ALERTA:
        resultados.append({
            "data": data_final.strftime("%d/%m/%Y"),
            "of": str(linha["OF"]),
            "status": status_original,
            "cliente": str(linha["Cliente"]),
            "setor": str(linha["Razão Social"])
        })

    print("🔎 TOTAL DE RESULTADOS:", len(resultados))

    if len(resultados) > 0:
        enviar_email_brevo(resultados)
    else:
        print("ℹ️ Nenhuma correspondência encontrada. E-mail não enviado.")


# =========================
# EXECUÇÃO
# =========================
if __name__ == "__main__":
    rodar_verificacao()
