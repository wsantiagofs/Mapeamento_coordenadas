import timeit
import geocoder
import pandas as pd
import time

# Lista de provedores de geocodificação a serem usados
geocoders = [geocoder.osm, geocoder.arcgis, geocoder.google]

# Função para obter as coordenadas de um endereço usando múltiplos provedores
def get_geocode(location):
    for geocoder_func in geocoders:
        g = geocoder_func(location)
        if g.ok:
            return g.latlng
        else:
            print(f"Falha ao obter coordenadas para {location} usando {geocoder_func.__name__}. Mensagem de erro: {g.status}")
    print(f'Não foi possível encontrar as coordenadas para {location}.')
    return None

# Carrega o arquivo Excel em um DataFrame pandas
df = pd.read_excel('Codigo_coordenadas.xlsx')

# Cria uma nova coluna que é a concatenação das colunas, com uma vírgula como separador
df['Endereço'] = df['Rua'].astype(str) + ', ' + df['Municipio'].astype(str) + ', ' + df['Bairro'].astype(str) + ', ' + df['Estado'].astype(str)

# Loop para buscar as coordenadas e salvar no DataFrame
for idx, row in df.iterrows():
    coords = get_geocode(row['Endereço'])
    if coords:
        df.at[idx, 'Latitude'] = coords[0]
        df.at[idx, 'Longitude'] = coords[1]
    time.sleep(1)  # Adiciona um atraso de 1 segundo entre as solicitações para evitar limites de taxa

# Salva o DataFrame atualizado de volta para o arquivo Excel
df.to_excel('Codigo_coordenadas_atualizado.xlsx', index=False)

# Inicio do temporizador
start = timeit.default_timer()
# Final do temporizador
end = timeit.default_timer()
print(f'Tempo de execução: {end - start} segundos')

# Inicio do temporizador
start = timeit.default_timer()

# Carrega o arquivo Excel em um DataFrame pandas
tabela = pd.read_excel('C:\\Users\\gabrielle.fuentes\\OneDrive - DEEPESSOAS\\Documentos\\Copasa\\Codigo_para_coordenadas\\Codigo_coordenadas.xlsx')

# Cria uma nova coluna que é a concatenação das colunas, com uma vírgula como separador
tabela['Endereço'] = tabela['Rua'].astype(str) + ', ' + tabela['Bairro'].astype(str) + ', ' + tabela['Municipio'].astype(str)

# Função para obter as coordenadas
def get_geocode(location):
    g = geocoder.osm(location)
    if g.ok:
        return g.latlng
    else:
        print('Não foi possível encontrar as coordenadas para a localização fornecida.')
        return None, None

# Loop para buscar as coordenadas e salvar no DataFrame
for idx, row in tabela.iterrows():
    coords = get_geocode(row['Endereço'])
    if coords:
        tabela.at[idx, 'Latitude'] = coords[0]
        tabela.at[idx, 'Longitude'] = coords[1]

# Salva o DataFrame atualizado de volta para o arquivo Excel
tabela.to_excel('C:\\Users\\gabrielle.fuentes\\OneDrive - DEEPESSOAS\\Documentos\\Copasa\\Codigo_para_coordenadas\\Codigo_coordenadas.xlsx', index=False)

# Final do temporizador
end = timeit.default_timer()
print(f'Tempo de execução: {end - start} segundos')
