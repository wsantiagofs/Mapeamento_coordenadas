import pandas as pd
import geopandas
from shapely.geometry import Point
import json # Necessário se o seu geojson_file ainda for gerado ou lido como string

# --- 1. Definir os caminhos dos arquivos ---
excel_file = 'teste2.xlsx'  # Nome do seu arquivo Excel
geojson_file = 'MunicipioGuaruja.geojson' # Nome do seu arquivo GeoJSON (verifique o nome real do seu arquivo!)

# --- 4. Leitura dos dados do Excel ---
try:
    df_dados = pd.read_excel(excel_file)
    print("\nDados do Excel lidos com sucesso (nomes originais das colunas):")
    print(df_dados.head())
    print("Nomes originais das colunas lidas:", df_dados.columns.tolist())

    # --- NOVO BLOCO: RENOMEAR COLUNAS PELO ÍNDICE ---
    # Mapeamento dos índices para os novos nomes das colunas
    # CERTIFIQUE-SE QUE ESSES ÍNDICES (0, 3, 8, 9) CORRESPONDEM EXATAMENTE ÀS COLUNAS
    # MATRICULA, COMUNIDADE, Latitude e Longitude NO SEU ARQUIVO EXCEL.
    # Se suas colunas de Latitude/Longitude mudarem de posição ou forem outras, ajuste os índices.
    col_mapping = {
        df_dados.columns[0]: 'MATRICULA_PADRONIZADA', # Coluna A (índice 0)
        df_dados.columns[3]: 'COMUNIDADE_PADRONIZADA', # Coluna D (índice 3)
        df_dados.columns[8]: 'Latitude_PADRONIZADA',  # Coluna I (índice 8)
        df_dados.columns[9]: 'Longitude_PADRONIZADA' # Coluna J (índice 9)
    }

    # Renomear as colunas no DataFrame
    df_dados.rename(columns=col_mapping, inplace=True)
    print("\nColunas após renomeio por índice para padronização:")
    print(df_dados.columns.tolist())

except FileNotFoundError:
    print(f"Erro: Arquivo Excel '{excel_file}' não encontrado. Verifique o caminho e o nome do arquivo.")
    exit()

# --- TRATAMENTO DAS COLUNAS DE LATITUDE E LONGITUDE (usando os nomes padronizados) ---
# Substituir vírgula por ponto e converter para float
df_dados['Latitude_PADRONIZADA'] = df_dados['Latitude_PADRONIZADA'].astype(str).str.replace(',', '.', regex=False).astype(float)
df_dados['Longitude_PADRONIZADA'] = df_dados['Longitude_PADRONIZADA'].astype(str).str.replace(',', '.', regex=False).astype(float)

print("\nDados do Excel após tratamento e conversão para float:")
print(df_dados.head())
print(df_dados.info())

# --- 5. Criação do GeoDataFrame a partir dos pontos do Excel (usando os nomes padronizados) ---
geometry = [Point(xy) for xy in zip(df_dados['Longitude_PADRONIZADA'], df_dados['Latitude_PADRONIZADA'])]
gdf_pontos = geopandas.GeoDataFrame(df_dados, geometry=geometry, crs="EPSG:4326")

print("\nGeoDataFrame de pontos criado:")
print(gdf_pontos.head())

# --- 6. Leitura do arquivo GeoJSON (área de atuação) ---
try:
    gdf_area_guaruja = geopandas.read_file(geojson_file)
    # Garante que o CRS (Sistema de Referência de Coordenadas) seja o mesmo para comparação
    gdf_area_guaruja = gdf_area_guaruja.to_crs("EPSG:4326")
    print("\nGeoJSON da área de atuação do Guarujá lido com sucesso:")
    print(gdf_area_guaruja)
except FileNotFoundError:
    print(f"Erro: Arquivo GeoJSON '{geojson_file}' não encontrado. Verifique o caminho e o nome do arquivo.")
    exit()
except Exception as e:
    print(f"Erro ao ler o arquivo GeoJSON '{geojson_file}': {e}")
    print("Verifique se o arquivo GeoJSON está bem formatado e não está corrompido.")
    exit()

# --- 7. Comparação Espacial: Verificar se os pontos estão dentro da área do Guarujá ---
# Combina todas as geometrias em uma única (se o GeoJSON tiver mais de um polígono)
area_guaruja_polygon = gdf_area_guaruja.geometry.unary_union

# Adiciona uma nova coluna ao GeoDataFrame de pontos indicando se está dentro da área
gdf_pontos['Dentro_Area_Guaruja'] = gdf_pontos.geometry.apply(lambda point: point.within(area_guaruja_polygon))

print("\nResultados da comparação espacial:")
# Usar os novos nomes padronizados no print
print(gdf_pontos[['MATRICULA_PADRONIZADA', 'COMUNIDADE_PADRONIZADA', 'Latitude_PADRONIZADA', 'Longitude_PADRONIZADA', 'Dentro_Area_Guaruja']])

# --- Opcional: Filtrar pontos que estão fora da área ---
pontos_fora = gdf_pontos[gdf_pontos['Dentro_Area_Guaruja'] == False]
if not pontos_fora.empty:
    print("\nPontos encontrados FORA da área de atuação do Guarujá:")
    print(pontos_fora[['MATRICULA_PADRONIZADA', 'COMUNIDADE_PADRONIZADA', 'Latitude_PADRONIZADA', 'Longitude_PADRONIZADA']])
else:
    print("\nTodos os pontos estão dentro da área de atuação do Guarujá.")

# --- Próximo passo: Salvar os dados corrigidos ou marcados ---
# Você pode salvar este GeoDataFrame de volta para um CSV ou Excel,
# incluindo a nova coluna 'Dentro_Area_Guaruja'.
gdf_pontos.to_excel('seus_dados_verificados.xlsx', index=False)
print("\nDados com a coluna 'Dentro_Area_Guaruja' salvos em 'seus_dados_verificados.xlsx'.")