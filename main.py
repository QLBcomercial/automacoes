import os
import requests
import pandas as pd
from datetime import datetime, timedelta

# =========================
# CONFIGURAÇÕES
# =========================
SHEET_ID = "1A0beFGh1PL-t7PTuZvRRuuk-nDQeWZxsMPVQ1I4QM0I"
SHEET_NAME = "Pedidos"
URL_PLANILHA = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/gviz/tq?tqx=out:csv&sheet={SHEET_NAME}"

# =========================
# FUNÇÕES AUXILIARES
# =========================
def soma_dias_uteis(data_inicial, dias):
    data = data_inicial
    adicionados = 0
    while adicionados < dias:
        data += timedelta(days=1)
        if data.weekday() < 5:  # 0=segunda, 4=sexta
            adicionados += 1
    return data


def interpretar_data(valor):
    try:
        texto = str(valor)
        if "à" in texto:
            texto = texto.split("à")[-1]
        texto = texto.strip()
        return datetime.strptime(texto, "%d/%m/%Y")
    except:
        return None


def normalizar_status(status):
    return (
        str(status)
        .strip()
        .lower()
        .replace("ç", "c")
        .replace("ã", "a")
        .replace("á", "a")
        .replace("é", "e")
        .replace("í", "i")
        .replace("ó", "o")
        .replace("ú", "u")
    )

# =========================
# ROTINA PRINCIPAL
# =========================
def rodar_verificacao():
    print("📥 Lendo planilha...")
    df = pd.read_csv(URL_PLANILHA).fillna("")

    hoje = datetime.now()
    limite = soma_dias_uteis(hoje, 3)

    resultados = []

    for _, linha in df.iterrows():
        data_texto = linha.iloc[0]
        of_valor = linha.iloc[1]
        status_original = linha.iloc[2]
        cliente = linha.iloc[3]
        cliente_a = linha.iloc[8]

        status = normalizar_status(status_original)
        data_linha = interpretar_data(data_texto)

        # Somente estes status entram
        if "producao" not in status and "nova" not in status:
    continue

        if not data_linha:
            continue

        # Define tipo de alerta
        if data_linha < hoje:
            tipo = "ATRASADO"
        elif hoje <= data_linha <= limite:
            tipo = "PRÓXIMO DO VENCIMENTO"
        else:
            continue

        # Limpeza da OF
        try:
            of_limpa = str(int(float(of_valor)))
        except:
            of_limpa = str(of_valor)

        resultados.append({
            "data": data_linha.strftime("%d/%m/%Y"),
            "of": of_limpa,
            "status": status_original,
            "tipo": tipo,
            "cliente": cliente,
            "cliente_a": cliente_a
        })

    print("🚨 DEBUG: quantidade de resultados =", len(resultados))
    print("🚨 DEBUG: chamando função de envio")
    enviar_email_brevo(resultados)



# =========================
# ENVIO DE E-MAIL (BREVO)
# =========================
def enviar_email_brevo(dados):
    api_key = os.getenv("BREVO_API_KEY", "").strip()
    if not api_key:
        print("❌ ERRO: BREVO_API_KEY não encontrada.")
        return

    url = "https://api.brevo.com/v3/smtp/email"
    headers = {
        "api-key": api_key,
        "Content-Type": "application/json",
        "accept": "application/json"
    }

    corpo_tabela = ""
    for d in dados:
        cor = "#ffcccc" if d["tipo"] == "ATRASADO" else "#fff2cc"
        corpo_tabela += f"""
        <tr style="background-color:{cor}">
            <td>{d['data']}</td>
            <td>{d['of']}</td>
            <td>{d['status']}</td>
            <td>{d['tipo']}</td>
            <td>{d['cliente']}</td>
            <td>{d['cliente_a']}</td>
        </tr>
        """

    html_content = f"""
    <html>
    <body>
        <h3>⚠️ Relatório de OFs – Pendências</h3>
        <table border="1" cellpadding="5" cellspacing="0" width="100%">
            <tr style="background-color:#eaeaea">
                <th>Data</th>
                <th>OF</th>
                <th>Status</th>
                <th>Tipo</th>
                <th>Cliente</th>
                <th>Cliente A</th>
            </tr>
            {corpo_tabela}
        </table>
    </body>
    </html>
    """

    payload = {
        "sender": {"name": "Sistema Quimlab", "email": "quimlabcomercial@gmail.com"},
        "to": [
            {"email": "marcos@quimlab.com.br"},
            {"email": "rodrigo@quimlab.com.br"}
        ],
        "subject": "⚠️ Relatório de OFs – Atrasos e Alertas",
        "htmlContent": html_content
    }

    response = requests.post(url, headers=headers, json=payload)
    
    if resultados:
        print(f"✅ {len(resultados)} pendência(s) encontrada(s). Enviando e-mail...")
        enviar_email_brevo(resultados)
    else:
        print("ℹ️ Nenhuma pendência encontrada. E-mail não enviado.")



    
    if response.status_code != 201:
        print(response.text)


# =========================
# EXECUÇÃO
# =========================
if __name__ == "__main__":
    rodar_verificacao()
