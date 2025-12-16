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
        # Pega a última data se for um intervalo
        parte_final = str(valor).split('à')[-1].split('a')[-1].strip()
        return datetime.strptime(parte_final, "%d/%m/%Y")
    except:
        return None

def rodar_verificacao():
    print("Iniciando leitura da planilha...")
    # O .fillna('') substitui todos os campos vazios por nada, removendo o 'nan'
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
                resultados.append({
                    "data": data_linha.strftime("%d/%m/%Y"),
                    "of": linha.iloc[1],
                    "status": status,
                    "cliente": linha.iloc[3],
                    "cliente_a": linha.iloc[8] # Agora virá vazio se não houver dados
                })

    if resultados:
        print(f"Sucesso: {len(resultados)} pendências encontradas.")
        enviar_email_brevo(resultados)
    else:
        print("Nenhuma pendência encontrada.")

def enviar_email_brevo(dados):
    if not BREVO_API_KEY:
        print("ERRO CRÍTICO: BREVO_API_KEY não configurada no GitHub Secrets.")
        return

    url = "https://api.brevo.com/v3/smtp/email"
    headers = {"api-key": str(os.getenv("BREVO_API_KEY")).strip(), "Content-Type": "application/json"}
    
    # Montagem da tabela de forma mais limpa
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

    payload = {
        "sender": {"name": "Sistema Quimlab", "email": "quimlabcomercial@gmail.com"},
        "to": [{"email": "marcos@quimlab.com.br"}, {"email": "rodrigo@quimlab.com.br"}],
        "subject": "⚠️ Relatório de OFs em Atraso",
        "htmlContent": html_content
    }

    response = requests.post(url, headers=headers, json=payload)
    print(f"Status Brevo: {response.status_code}")
    if response.status_code != 201:
        print(f"Erro: {response.text}")

if __name__ == "__main__":
    rodar_verificacao()
