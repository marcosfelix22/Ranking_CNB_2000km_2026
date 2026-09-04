import os
import requests
import pandas as pd

# --- CONFIGURAÇÕES ---
CLIENT_ID = os.environ.get('CLIENT_ID')
CLIENT_SECRET = os.environ.get('CLIENT_SECRET')
NOME_ARQUIVO = 'Ranking_CNB_2026.xlsx'

# Dicionário com os atletas e os tokens cadastrados nos Secrets
ATLETAS = {
    "Marcos Felix": os.environ.get('TOKEN_MARCOS'),
    # Para novos atletas, basta adicionar a linha abaixo:
    # "Nome do Atleta": os.environ.get('TOKEN_NOME_DO_ATLETA'),
}

def obter_access_token(refresh_token):
    if not refresh_token:
        print("❌ Erro: refresh_token veio vazio ou não foi lido do Secret.")
        return None
    
    payload = {
        'client_id': CLIENT_ID,
        'client_secret': CLIENT_SECRET,
        'refresh_token': refresh_token,
        'grant_type': 'refresh_token'
    }
    
    try:
        res = requests.post("https://www.strava.com/oauth/token", data=payload)
        print(f"--- DIAGNÓSTICO STRAVA ---")
        print(f"Status HTTP: {res.status_code}")
        print(f"Resposta Completa: {res.text}")
        print(f"---------------------------")
        
        if res.status_code == 200:
            return res.json().get('access_token')
        else:
            return None
    except Exception as e:
        print(f"Erro na requisição: {e}")
        return None

def formatar_km(valor):
    return f"{valor:,.1f}".replace(",", "X").replace(".", ",").replace("X", ".") + " km"

def formatar_alt(valor):
    return f"{int(valor):,}".replace(",", ".") + " m"

dados_ranking = []

# Processa atleta por atleta
for nome_atleta, ref_token in ATLETAS.items():
    if not ref_token:
        print(f"Aviso: Secret para '{nome_atleta}' não encontrado.")
        continue

    access_token = obter_access_token(ref_token)
    if not access_token:
        print(f"Erro: Falha na renovação do token para '{nome_atleta}'.")
        continue

    headers = {'Authorization': f'Bearer {access_token}'}
    url = "https://www.strava.com/api/v3/athlete/activities"
    
    resposta = requests.get(url, headers=headers, params={'per_page': 200, 'page': 1})
    
    if resposta.status_code == 200:
        atividades = resposta.json()
        km_total = 0.0
        alt_total = 0.0
        treinos = 0
        
        for act in atividades:
            tipo = act.get('type')
            data_inicio = act.get('start_date', '')
            
            # Filtra corridas (Run / TrailRun) realizadas em 2026
            if tipo in ['Run', 'TrailRun'] and data_inicio.startswith('2026'):
                dist_km = act.get('distance', 0.0) / 1000.0
                alt = act.get('total_elevation_gain', 0.0)
                
                km_total += dist_km
                alt_total += alt
                treinos += 1

        dados_ranking.append({
            'Atleta': nome_atleta,
            'KM Total Bruto': km_total,
            'KM Total': formatar_km(km_total),
            'Altimetria (m)': formatar_alt(alt_total),
            'Treinos': treinos
        })
        print(f"✓ {nome_atleta}: {formatar_km(km_total)} ({treinos} treinos)")
    else:
        print(f"Erro na consulta de {nome_atleta}: Status {resposta.status_code}")

# Organiza os dados e gera o ranking
if dados_ranking:
    df = pd.DataFrame(dados_ranking)
    df = df.sort_values(by='KM Total Bruto', ascending=False)
    df = df.drop(columns=['KM Total Bruto'])
else:
    df = pd.DataFrame(columns=['Atleta', 'KM Total', 'Altimetria (m)', 'Treinos'])

# Salva na nova planilha
with pd.ExcelWriter(NOME_ARQUIVO, engine='openpyxl') as writer:
    df.to_excel(writer, sheet_name='Ranking', index=False)

print(f"Planilha {NOME_ARQUIVO} sincronizada com sucesso!")
