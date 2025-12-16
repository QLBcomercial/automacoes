import os
import requests
import pandas as pd
from datetime import datetime, timedelta

# Configurações
SHEET_ID = "1A0beFGh1PL-t7PTuZvRRuuk-nDQeWZxsMPVQ1I4QM0I"
SHEET_NAME = "Pedidos"
RESEND_API_KEY = os.getenv("RESEND_API_KEY")
URL_PLANILHA = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/gviz/tq?tqx=out:csv&sheet={SHEET_NAME}"

def soma_dias_uteis(data_inicial, dias):
    data = data_inicial
    adicionados = 0
    while adicionados < dias:
        data += timedelta(days=1)
        if data.weekday() < 5:  # 0-4 são dias úteis (seg-sex)
            adicionados += 1
    return data

def interpretar_data(valor):
    try:
        # Se for intervalo "data à data", pega a última
        parte_final = str(valor).split('à')[-1].split('a')[-1].strip()
        return datetime.strptime(parte_final, "%d/%m/%Y")
    except:
        return None

def rodar_verificacao():
    # 1. Lê a planilha (pública para quem tem o link ou via Service Account)
    df = pd.read_csv(URL_PLANILHA)
    
    hoje = datetime.now()
    limite = soma_dias_uteis(hoje, 3)
    resultados = []

    # 2. Lógica de Filtragem
    # Coluna 0: Data, Coluna 1: OF, Coluna 2: Status, Coluna 3: Cliente, Coluna 8: Cliente A
    for index, linha in df.iterrows():
        status = str(linha.iloc[2])
        if status not in ["Em Produção", "Nova"]:
            continue
            
        data_linha = interpretar_data(linha.iloc[0])
        if data_linha and data_linha <= limite:
            resultados.append({
                "data": data_linha.strftime("%d/%m/%Y"),
                "of": linha.iloc[1],
                "status": status,
                "cliente": linha.iloc[3],
                "cliente_a": linha.iloc[8]
            })

    if resultados:
        enviar_email(resultados)
    else:
        print("Nenhuma pendência encontrada.")

def enviar_email(dados):
    url = "https://api.resend.com/emails"
    headers = {
        "Authorization": f"Bearer {RESEND_API_KEY}",
        "Content-Type": "application/json"
    }
    
    linhas_tabela = "".join([
        f"<tr><td>{d['data']}</td><td>{d['of']}</td><td>{d['status']}</td><td>{d['cliente']}</td><td>{d['cliente_a']}</td></tr>"
        for d in dados
    ])

    html_body = f"""
    <h3>⚠️ OFs em Atraso</h3>
    <table border="1" style="border-collapse: collapse;">
        <tr style="background-color: #f2f2f2;">
            <th>Data</th><th>OF</th><th>Status</th><th>Cliente</th><th>Cliente A</th>
        </tr>
        {linhas_tabela}
    </table>
    """

    payload = {
        "from": "Sistema Quimlab <onboarding@resend.dev>",
        "to": ["marcos@quimlab.com.br"],
        "subject": "⚠️ Relatório de OFs em Atraso",
        "html": html_body
    }

    response = requests.post(url, headers=headers, json=payload)
    print(f"Status do envio: {response.status_code}, Resposta: {response.text}")

if __name__ == "__main__":
    rodar_verificacao()
