import requests
import pandas as pd
import os

# --- CONFIGURAÇÃO ---
CLIENT_ID = os.environ.get('CLIENT_ID')
CLIENT_SECRET = os.environ.get('CLIENT_SECRET')
REFRESH_TOKEN = os.environ.get('REFRESH_TOKEN')
CLUB_ID = '1976744' 
NOME_ARQUIVO = 'Ranking_CNB_2000km_2026.xlsx'

def obter_access_token():
    payload = {'client_id': CLIENT_ID, 'client_secret': CLIENT_SECRET, 'refresh_token': REFRESH_TOKEN, 'grant_type': 'refresh_token'}
    res = requests.post("https://www.strava.com/oauth/token", data=payload).json()
    return res.get('access_token')

def formatar_km(valor):
    return f"{valor:,.1f}".replace(",", "X").replace(".", ",").replace("X", ".") + " km"

def formatar_alt(valor):
    return f"{int(valor):,}".replace(",", ".") + " m"

# 1. Carregar planilha ou criar nova do zero
if os.path.exists(NOME_ARQUIVO):
    try:
        with pd.ExcelFile(NOME_ARQUIVO) as reader:
            df_ranking = pd.read_excel(reader, sheet_name='Ranking')
            
            # Garante que as colunas existem antes de limpar
            if 'Atleta' in df_ranking.columns:
                df_ranking['Atleta'] = df_ranking['Atleta'].astype(str).str.strip()
                
                if 'KM Total' in df_ranking.columns and df_ranking['KM Total'].dtype == object:
                    df_ranking['KM Total'] = df_ranking['KM Total'].str.replace(' km', '', regex=False).str.replace('.', '', regex=False).str.replace(',', '.', regex=False).astype(float)
                
                if 'Altimetria (m)' in df_ranking.columns and df_ranking['Altimetria (m)'].dtype == object:
                    df_ranking['Altimetria (m)'] = df_ranking['Altimetria (m)'].str.replace(' m', '', regex=False).str.replace('.', '', regex=False).astype(float)
                
                df_ranking = df_ranking.set_index('Atleta').groupby(level=0).sum()
            else:
                # Se não tem a coluna Atleta, recomeça a tabela
                df_ranking = pd.DataFrame(columns=['Atleta', 'KM Total', 'Altimetria (m)']).set_index('Atleta')

            df_historico = pd.read_excel(reader, sheet_name='IDs_Processados')
            ids_ja_somados = set(df_historico['id'].astype(str).tolist())
    except Exception as e:
        print(f"Erro ao ler planilha, criando nova: {e}")
        df_ranking = pd.DataFrame(columns=['Atleta', 'KM Total', 'Altimetria (m)']).set_index('Atleta')
        ids_ja_somados = set()
else:
    df_ranking = pd.DataFrame(columns=['Atleta', 'KM Total', 'Altimetria (m)']).set_index('Atleta')
    ids_ja_somados = set()

# 2. Puxar do Strava
access_token = obter_access_token()
if access_token:
    for pagina in range(1, 11):
        url = f"https://www.strava.com/api/v3/clubs/{CLUB_ID}/activities"
        atividades = requests.get(url, headers={'Authorization': f'Bearer {access_token}'}, params={'per_page': 200, 'page': pagina}).json()
        if not atividades or 'errors' in atividades or len(atividades) == 0: break
        for act in atividades:
            id_unico = f"{act.get('distance')}_{act.get('elapsed_time')}_{act.get('athlete', {}).get('lastname')}"
            if id_unico not in ids_ja_somados:
                nome = f"{act.get('athlete', {}).get('firstname', 'Atleta')} {act.get('athlete', {}).get('lastname', '')}".strip()
                dist_km = act.get('distance', 0) / 1000
                alt = act.get('total_elevation_gain', 0)
                if dist_km > 0:
                    if nome not in df_ranking.index: df_ranking.loc[nome] = [0.0, 0.0]
                    df_ranking.at[nome, 'KM Total'] += dist_km
                    df_ranking.at[nome, 'Altimetria (m)'] += alt
                    ids_ja_somados.add(id_unico)

# 3. Ordenar e Salvar
df_ranking = df_ranking.sort_values(by='KM Total', ascending=False)
df_visual = df_ranking.reset_index().copy()
df_visual['KM Total'] = df_visual['KM Total'].apply(formatar_km)
df_visual['Altimetria (m)'] = df_visual['Altimetria (m)'].apply(formatar_alt)

with pd.ExcelWriter(NOME_ARQUIVO) as writer:
    df_visual.to_excel(writer, sheet_name='Ranking', index=False)
    pd.DataFrame(list(ids_ja_somados), columns=['id']).to_excel(writer, sheet_name='IDs_Processados', index=False)
