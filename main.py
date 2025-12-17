import os
import requests
import pandas as pd
from datetime import datetime, timedelta

# Configurações
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
        parte_final = str(valor).split('à')[-1].split('a')[-1].strip()
        return datetime.strptime(parte_final, "%d/%m/%Y")
    except:
        return None

def rodar_verificacao():
    print("Iniciando leitura da planilha...")
    # .fillna('') remove o "nan" das células vazias
    df = pd.read_csv(URL_PLANILHA).fillna('')
    
    hoje = datetime.now()
    limite = soma_dias_uteis(hoje, 3)
    resultados = []

    for index, linha in df.iterrows():
        status = str(linha.iloc[2]).strip()
        data_texto = linha.iloc[0]
        
        if status in ["Em Produção", "Nova"]:
            data_linha = interpretar_data(data_texto)
            if data_linha and data_linha <= limite:
                # Ajuste da OF: Remove .0 convertendo para inteiro
                of_valor = linha.iloc[1]
                try:
                    # Se for float (ex: 123.0) ou string que termina em .0
                    of_limpa = str(int(float(of_valor)))
                except:
                    of_limpa = str(of_valor)

                resultados.append({
                    "data": data_linha.strftime("%d/%m/%Y"),
                    "of": of_limpa,
                    "status": status,
                    "cliente": linha.iloc[3],
                    "cliente_a": linha.iloc[8]
                })

    if resultados:
        print(f"Sucesso: {len(resultados)} pendências encontradas. Enviando e-mail...")
        enviar_email_brevo(resultados)
    else:
        print("Nenhuma pendência encontrada.")

def enviar_email_brevo(dados):
    api_key = str(os.getenv("BREVO_API_KEY", "")).strip()
    if not api_key:
        print("ERRO CRÍTICO: BREVO_API_KEY não encontrada.")
        return

    url = "https://api.brevo.com/v3/smtp/email"
    headers = {
        "api-key": api_key,
        "Content-Type": "application/json",
        "accept": "application/json"
    }
    
    corpo_tabela = ""
    for d in dados:
        corpo_tabela += f"<tr><td>{d['data']}</td><td>{d['of']}</td><td>{d['status']}</td><td>{d['cliente']}</td><td>{d['cliente_a']}</td></tr>"

    html_content = f"""
    <html><body>
        <h3>Relatório de OFs em Atraso</h3>
        <table border="1" style="border-collapse: collapse; width: 100%;">
            <tr style="background-color: #eee;">
                <th>Data</th><th>OF</th><th>Status</th><th>Cliente</th><th>Cliente A</th>
            </tr>
            {corpo_tabela}
        </table>
    </body></html>
    """

    # PAYLOAD CORRIGIDO: Chaves simples para os dicionários internos
    payload = {
        "sender": {"name": "Sistema Quimlab", "email": "quimlabcomercial@gmail.com"},
        "to": [
            {"email": "marcos@quimlab.com.br"},
            {"email": "rodrigo@quimlab.com.br"}
        ],
        "subject": "⚠️ Relatório de OFs em Atraso",
        "htmlContent": html_content
    }

    response = requests.post(url, headers=headers, json=payload)
    print(f"Status Brevo: {response.status_code}")
    if response.status_code != 201:
        print(f"Detalhes: {response.text}")

if __name__ == "__main__":
    rodar_verificacao()
