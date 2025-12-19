import pandas as pd
import requests
import os
from datetime import datetime, timedelta
import unicodedata


# =========================
# CONFIGURAÇÕES
# =========================
SHEET_ID = "1A0beFGh1PL-t7PTuZvRRuuk-nDQeWZxsMPVQ1I4QM0I"
SHEET_NAME = "Pedidos"

URL_PLANILHA = (
    f"https://docs.google.com/spreadsheets/d/{SHEET_ID}"
    f"/gviz/tq?tqx=out:csv&sheet={SHEET_NAME}"
)

BREVO_API_KEY = os.getenv("BREVO_API_KEY")

EMAIL_REMETENTE = {
    "name": "Quimlab",
    "email": "noreply@seudominio.com.br"  # precisa estar VALIDADO na Brevo
}

EMAIL_DESTINATARIO = {
    "email": "destinatario@seudominio.com.br",
    "name": "Destinatário"
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


def interpretar_data(valor):
    if pd.isna(valor):
        return None

    texto = str(valor)

    # Caso venha no formato: 10/12/2025 à 16/12/2025
    if "à" in texto:
        texto = texto.split("à")[0].strip()

    try:
        return pd.to_datetime(texto, dayfirst=True).date()
    except Exception:
        return None


# =========================
# ENVIO DE EMAIL (BREVO)
# =========================
def enviar_email_brevo(resultados):
    print("📨 Função enviar_email_brevo chamada")

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
        <p>As seguintes Ordens de Fabricação requerem atenção:</p>
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
        "accept": "application/json",
        "content-type": "application/json",
        "api-key": BREVO_API_KEY
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
    limite = hoje + timedelta(days=DIAS_ALERTA)

    df = pd.read_excel(ARQUIVO_PLANILHA)

    print("📥 Planilha carregada")
    print("🔎 Total de linhas:", len(df))

    resultados = []

    for i, linha in df.iterrows():
        data_linha = interpretar_data(linha.iloc[0])
        if not data_linha:
            continue

        status_original = str(linha.iloc[2])
        status = normalizar_texto(status_original)

        # Filtro de status (flexível)
        if not any(p in status for p in ["produc", "nova"]):
            continue

        if data_linha < hoje:
            tipo = "ATRASADO"
        elif hoje <= data_linha <= limite:
            tipo = "PRÓXIMO DO PRAZO"
        else:
            continue

        resultados.append({
            "data": data_linha.strftime("%d/%m/%Y"),
            "of": str(linha.iloc[1]),
            "status": status_original,
            "cliente": str(linha.iloc[3]),
            "setor": str(linha.iloc[5])
        })

    print("🔎 TOTAL DE RESULTADOS:", len(resultados))

    # ✅ CONDIÇÃO EXPLÍCITA (SEM AMBIGUIDADE)
    if len(resultados) > 0:
        print("✅ Correspondências encontradas. Enviando e-mail...")
        enviar_email_brevo(resultados)
    else:
        print("ℹ️ Nenhuma correspondência válida encontrada. E-mail não enviado.")


# =========================
# EXECUÇÃO
# =========================
if __name__ == "__main__":
    rodar_verificacao()
