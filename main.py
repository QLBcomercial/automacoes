import os
import requests
import pandas as pd
from datetime import datetime, timedelta

# Configurações de acesso
SHEET_ID = "1A0beFGh1PL-t7PTuZvRRuuk-nDQeWZxsMPVQ1I4QM0I"
SHEET_NAME = "Pedidos"
BREVO_API_KEY = os.getenv("BREVO_API_KEY")
URL_PLANILHA = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/gviz/tq?tqx=out:csv&sheet={SHEET_NAME}"

def soma_dias_uteis(data_inicial, dias):
    data = data_inicial
    adicionados = 0
    while adicionados < dias:
        data += timedelta(days=1)
        if data.weekday() < 5: 
            adicionados += 1
    return data

def interpretar_data(valor):
    try:
        # Pega a última data se for um intervalo (ex: 10/12/2024 à 12/12/2024)
        parte_final = str(valor).split('à')[-1].split('a')[-1].strip()
        return datetime.strptime(parte_final, "%d/%m/%Y")
    except:
        return None

def rodar_verificacao():
    print("Iniciando leitura da planilha...")
    try:
        df = pd.read_csv(URL_PLANILHA)
    except Exception as e:
        print(f"Erro ao ler planilha: {e}")
        return

    hoje = datetime.now()
    limite = soma_dias_uteis(hoje, 3)
    resultados = []

    for index, linha in df.iterrows():
        status = str(linha.iloc[2]).strip()
        data_texto = linha.iloc[0]
        
        if status in ["Em Produção", "Nova"]:
            data_linha = interpretar_data(data_texto)
            if data_linha and data_linha <= limite:
                resultados.append({
                    "data": data_linha.strftime("%d/%m/%Y"),
                    "of": linha.iloc[1],
                    "status": status,
                    "cliente": linha.iloc[3],
                    "cliente_a": linha.iloc[8]
                })

    if resultados:
        print(f"Encontradas {len(resultados)} pendências. Enviando e-mail...")
        enviar_email_brevo(resultados)
    else:
        print("Nenhuma pendência encontrada com os critérios.")

def enviar_email_brevo(dados):
    url = "https://api.brevo.com/v3/smtp/email"
    
    # Validação da chave antes do envio
    if not BREVO_API_KEY:
        print("ERRO: A BREVO_API_KEY não foi encontrada nas variáveis de ambiente.")
        return

    headers = {
        "api-key": BREVO_API_KEY,
        "Content-Type": "application/json"
    }
    
    linhas_tabela = "".join([
        f"<tr><td>{d['data']}</td><td>{d['of']}</td><td>{d['status']}</td><td>{d['cliente']}</td><td>{d['cliente_a']}</td></tr>"
        for d in dados
    ])

    html_content = f"""
    <html>
    <body style="font-family: Arial, sans-serif;">
        <p>Olá,</p>
        <p>Seguem as OFs que estão em atraso ou próximas do limite:</p>
        <table border="1" style="border-collapse: collapse; width: 100%;">
            <tr style="background-color: #f2f2f2;">
                <th>Data</th><th>OF</th><th>Status</th><th>Cliente</th><th>Cliente A</th>
            </tr>
            {linhas_tabela}
        </table>
        <p><br>— Sistema Automático Quimlab</p>
    </body>
    </html>
    """

    payload = {
        "sender": {"name": "Sistema Quimlab", "email": "quimlabcomercial@gmail.com"},
        "to": [
            {"email": "marcos@quimlab.com.br"},
            {"email": "quimlabcomercial@gmail.com"}
        ],
        "subject": "⚠️ Relatório de OFs em Atraso",
        "htmlContent": html_content
    }

    response = requests.post(url, headers=headers, json=payload)
    print(f"Status Brevo: {response.status_code}")
    if response.status_code != 201:
        print(f"Detalhes do erro: {response.text}")

if __name__ == "__main__":
    rodar_verificacao()
