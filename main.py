import pandas as pd
import requests
import os
from datetime import datetime, timedelta
import unicodedata
import re


# =========================
# CONFIGURAÇÕES
# =========================
URL_PLANILHA = (
    "https://docs.google.com/spreadsheets/d/"
    "1A0beFGh1PL-t7PTuZvRRuuk-nDQeWZxsMPVQ1I4QM0I"
    "/gviz/tq?tqx=out:csv&sheet=Pedidos"
)

DIAS_ALERTA = 3
BREVO_API_KEY = os.getenv("BREVO_API_KEY")

EMAIL_REMETENTE = {
    "name": "Quimlab",
    "email": "quimlabcomercial@gmail.com"
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

    # encontra TODAS as datas no texto
    datas = re.findall(r"\d{1,2}/\d{1,2}/\d{4}", texto)

    if not datas:
        return None

    # a data final é sempre a última
    return pd.to_datetime(datas[-1], dayfirst=True).date()

def dias_uteis_entre(hoje, data_final):
    dias = 0
    data = hoje

    while data < data_final:
        data += timedelta(days=1)
        if data.weekday() < 5:
            dias += 1

    return dias


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
            <td>{r['cliente a']}</td>
        </tr>
        """

    html = f"""
    <html>
    <body>
        <p>Ordens de Fabricação com prazo crítico:</p>
        <table border="1" cellpadding="5" cellspacing="0">
            <tr>
                <th>Data Final</th>
                <th>OF</th>
                <th>Status</th>
                <th>Cliente</th>
                <th>Setor</th>
                <th>Cliente A</th>
            </tr>
            {linhas_html}
        </table>
    </body>
    </html>
    """

    payload = {
        "sender": EMAIL_REMETENTE,
        "to": [EMAIL_DESTINATARIO],
        "subject": "⚠️ Alerta – Ordens de Fabricação (≤ 3 dias úteis)",
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

    if not BREVO_API_KEY:
        print("❌ ERRO: BREVO_API_KEY não encontrada")
        return

    hoje = datetime.today().date()

    print("🌐 Lendo planilha Google Sheets (CSV)")
    df = pd.read_csv(URL_PLANILHA, dtype=str)
   
    print("📌 COLUNAS ENCONTRADAS NO CSV:")
    for col in df.columns:
        print(f"- [{col}]")


    print("📥 Linhas carregadas:", len(df))

    resultados = []

    for _, linha in df.iterrows():

        data_final = obter_data_final(
            linha.get("Data", linha.get("Data ", ""))
        )
        print(
            f"DEBUG | OF={linha.get('OF')} | "
            f"Data='{linha.get('Data')}' | "
            f"Status='{linha.get('Status')}'"
        )

        if not data_final:
            continue

        status_original = str(
            linha.get("Status", linha.get("Status ", ""))
        ).strip()

        status = normalizar_texto(status_original)

        # Apenas "Em Produção"
        if "produc" not in status:
            continue

        # ATRASADO
        if hoje > data_final:
            dias_uteis = 0
            tipo = "ATRASADO"
        else:
            dias_uteis = dias_uteis_entre(hoje, data_final)
            if dias_uteis > DIAS_ALERTA:
                continue
            tipo = "PRÓXIMO DO PRAZO"

        print(
            f"OF {linha['OF']} | {tipo} | Data final: {data_final} | Dias úteis: {dias_uteis}"
        )

        resultados.append({
            "data": data_final.strftime("%d/%m/%Y"),
            "of": str(linha["OF"]),
            "status": status_original,
            "cliente": str(linha["Cliente"]),
            "setor": str(linha["Razão Social"]),
            "cliente a": "" if pd.isna(linha["Cliente A"]) else str(linha["Cliente A"])
        })

    print("🔎 TOTAL DE RESULTADOS:", len(resultados))

    if resultados:
        enviar_email_brevo(resultados)
    else:
        print("ℹ️ Nenhuma correspondência encontrada. E-mail não enviado.")


# =========================
# EXECUÇÃO
# =========================
if __name__ == "__main__":
    rodar_verificacao()
