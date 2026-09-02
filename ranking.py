import os
import requests
import pandas as pd

# --- CONFIGURAÇÃO ---
CLIENT_ID = os.environ.get('CLIENT_ID')
CLIENT_SECRET = os.environ.get('CLIENT_SECRET')
REFRESH_TOKEN = os.environ.get('REFRESH_TOKEN')
CLUB_ID = '1976744' 
NOME_ARQUIVO = 'Ranking_CNB_2000km_2026.xlsx'

def obter_access_token():
    payload = {
        'client_id': CLIENT_ID, 
        'client_secret': CLIENT_SECRET, 
        'refresh_token': REFRESH_TOKEN, 
        'grant_type': 'refresh_token'
    }
    res = requests.post("https://www.strava.com/oauth/token", data=payload).json()
    return res.get('access_token')

def formatar_km(valor):
    return f"{valor:,.1f}".replace(",", "X").replace(".", ",").replace("X", ".") + " km"

def formatar_alt(valor):
    return f"{int(valor):,}".replace(",", ".") + " m"

# 1. Carregar dados existentes ou iniciar do zero
df_ranking = pd.DataFrame(columns=['KM Total', 'Altimetria (m)', 'Treinos'])
df_ranking.index.name = 'Atleta'
ids_ja_somados = set()

if os.path.exists(NOME_ARQUIVO):
    try:
        with pd.ExcelFile(NOME_ARQUIVO) as reader:
            if 'Ranking' in reader.sheet_names:
                df_temp = pd.read_excel(reader, sheet_name='Ranking')
                if 'Atleta' in df_temp.columns:
                    # Limpa strings de formatação anteriores se existirem
                    if 'KM Total' in df_temp.columns:
                        df_temp['KM Total'] = df_temp['KM Total'].astype(str).str.replace(' km', '', regex=False)\
                                                                .str.replace('.', '', regex=False)\
                                                                .str.replace(',', '.', regex=False)
                        df_temp['KM Total'] = pd.to_numeric(df_temp['KM Total'], errors='coerce').fillna(0.0)

                    if 'Altimetria (m)' in df_temp.columns:
                        df_temp['Altimetria (m)'] = df_temp['Altimetria (m)'].astype(str).str.replace(' m', '', regex=False)\
                                                                            .str.replace('.', '', regex=False)
                        df_temp['Altimetria (m)'] = pd.to_numeric(df_temp['Altimetria (m)'], errors='coerce').fillna(0.0)

                    if 'Treinos' in df_temp.columns:
                        df_temp['Treinos'] = pd.to_numeric(df_temp['Treinos'], errors='coerce').fillna(0).astype(int)
                    else:
                        df_temp['Treinos'] = 0

                    df_ranking = df_temp.set_index('Atleta')[['KM Total', 'Altimetria (m)', 'Treinos']]

            if 'IDs_Processados' in reader.sheet_names:
                df_historico = pd.read_excel(reader, sheet_name='IDs_Processados')
                if 'id' in df_historico.columns:
                    ids_ja_somados = set(df_historico['id'].astype(str).tolist())
    except Exception as e:
        print(f"Aviso ao ler planilha antiga (recriando estrutura): {e}")

# 2. Puxar novas atividades do Strava
access_token = obter_access_token()

if access_token:
    for pagina in range(1, 11):
        url = f"https://www.strava.com/api/v3/clubs/{CLUB_ID}/activities"
        resposta = requests.get(
            url, 
            headers={'Authorization': f'Bearer {access_token}'}, 
            params={'per_page': 200, 'page': pagina}
        )
        
        if resposta.status_code != 200:
            print(f"Erro ao acessar API do Strava: Status {resposta.status_code}")
            break
            
        atividades = resposta.json()
        if not atividades or isinstance(atividades, dict) and 'errors' in atividades:
            break

        for act in atividades:
            # Cria ID único confiável para a atividade
            dist = act.get('distance', 0)
            elapsed = act.get('elapsed_time', 0)
            lastname = act.get('athlete', {}).get('lastname', '')
            id_unico = f"{dist}_{elapsed}_{lastname}"

            if id_unico not in ids_ja_somados:
                firstname = act.get('athlete', {}).get('firstname', 'Atleta')
                nome = f"{firstname} {lastname}".strip()
                dist_km = dist / 1000.0
                alt = act.get('total_elevation_gain', 0.0)

                if dist_km > 0:
                    if nome not in df_ranking.index:
                        df_ranking.loc[nome] = [0.0, 0.0, 0]

                    df_ranking.at[nome, 'KM Total'] += dist_km
                    df_ranking.at[nome, 'Altimetria (m)'] += alt
                    df_ranking.at[nome, 'Treinos'] = int(df_ranking.at[nome, 'Treinos']) + 1
                    ids_ja_somados.add(id_unico)
else:
    print("Erro: Não foi possível obter o access_token do Strava.")

# 3. Agrupar, Ordenar e Salvar
df_ranking = df_ranking.groupby(level=0).sum()
df_ranking = df_ranking.sort_values(by='KM Total', ascending=False)

df_visual = df_ranking.reset_index().copy()
df_visual['KM Total'] = df_visual['KM Total'].apply(formatar_km)
df_visual['Altimetria (m)'] = df_visual['Altimetria (m)'].apply(formatar_alt)
df_visual['Treinos'] = df_visual['Treinos'].astype(int)

# Salva o arquivo final
with pd.ExcelWriter(NOME_ARQUIVO, engine='openpyxl') as writer:
    df_visual.to_excel(writer, sheet_name='Ranking', index=False)
    pd.DataFrame(list(ids_ja_somados), columns=['id']).to_excel(writer, sheet_name='IDs_Processados', index=False)

print("Sincronização de 2000km concluída com sucesso!")
