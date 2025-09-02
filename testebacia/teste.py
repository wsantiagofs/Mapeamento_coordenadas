import pandas as pd
import geopandas
from shapely.geometry import Point
import folium

# --- 1. Definir os caminhos dos arquivos ---
excel_file = 'teste2.xlsx'  # Nome do seu arquivo Excel
geojson_file = 'MunicipioGuaruja.geojson' # Nome do seu arquivo GeoJSON

# --- 4. Leitura dos dados do Excel ---
try:
    df_dados = pd.read_excel(excel_file)
    # ... (código para renomear colunas - mantenha o código existente)
    col_mapping = {
        df_dados.columns[0]: 'MATRICULA_PADRONIZADA',
        df_dados.columns[3]: 'COMUNIDADE_PADRONIZADA',
        df_dados.columns[8]: 'Latitude_PADRONIZADA',
        df_dados.columns[9]: 'Longitude_PADRONIZADA'
    }
    df_dados.rename(columns=col_mapping, inplace=True)

    # --- TRATAMENTO DAS COLUNAS DE LATITUDE E LONGITUDE ---
    # Passo 1: Garantir que são strings e substituir vírgulas por pontos
    df_dados['Latitude_PADRONIZADA'] = df_dados['Latitude_PADRONIZADA'].astype(str).str.replace(',', '.', regex=False)
    df_dados['Longitude_PADRONIZADA'] = df_dados['Longitude_PADRONIZADA'].astype(str).str.replace(',', '.', regex=False)

    # Passo 2: Converter para numérico. Qualquer valor que não possa ser convertido vira NaN.
    df_dados['Latitude_PADRONIZADA'] = pd.to_numeric(df_dados['Latitude_PADRONIZADA'], errors='coerce')
    df_dados['Longitude_PADRONIZADA'] = pd.to_numeric(df_dados['Longitude_PADRONIZADA'], errors='coerce')

    print("\nDados do Excel após tratamento e conversão para float:")
    print(df_dados.head())
    print(df_dados.info()) # Verifique os dtypes e contagem de non-null aqui para ver os NaNs

    # --- REMOVER LINHAS COM NaN NAS COORDENADAS ---
    # Esta linha vai agora remover os NaNs que foram gerados pelo errors='coerce'
    df_dados_validos = df_dados.dropna(subset=['Latitude_PADRONIZADA', 'Longitude_PADRONIZADA'])
    print(f"\nRemovidas {len(df_dados) - len(df_dados_validos)} linhas com coordenadas ausentes/inválidas.")
    print(f"Total de linhas válidas para mapeamento (após remover NaNs): {len(df_dados_validos)}")


    # --- 5. Criação do GeoDataFrame a partir dos pontos válidos (sem NaN) ---
    geometry = [Point(xy) for xy in zip(df_dados_validos['Longitude_PADRONIZADA'], df_dados_validos['Latitude_PADRONIZADA'])]
    gdf_pontos = geopandas.GeoDataFrame(df_dados_validos, geometry=geometry, crs="EPSG:4326")

    # ... (restante do seu código para ler GeoJSON, comparar e gerar mapa)
    # 6. Leitura do arquivo GeoJSON (área de atuação)
    gdf_area_guaruja = geopandas.read_file(geojson_file)
    gdf_area_guaruja = gdf_area_guaruja.to_crs("EPSG:4326")
    print("\nGeoJSON da área de atuação do Guarujá lido com sucesso:")
    print(gdf_area_guaruja)

    # 7. Comparação Espacial: Verificar se os pontos estão dentro da área do Guarujá
    area_guaruja_polygon = gdf_area_guaruja.geometry.unary_union
    gdf_pontos['Dentro_Area_Guaruja'] = gdf_pontos.geometry.apply(lambda point: point.within(area_guaruja_polygon))

    # NOVO BLOCO: Filtrar apenas os pontos que estão DENTRO da bacia delimitada
    gdf_pontos_filtrados = gdf_pontos[gdf_pontos['Dentro_Area_Guaruja'] == True].copy()
    print(f"\nTotal de pontos que estão DENTRO da bacia delimitada (e sem NaNs): {len(gdf_pontos_filtrados)}")
    print("Detalhes dos pontos dentro da bacia:")
    print(gdf_pontos_filtrados[['MATRICULA_PADRONIZADA', 'COMUNIDADE_PADRONIZADA', 'Latitude_PADRONIZADA', 'Longitude_PADRONIZADA', 'Dentro_Area_Guaruja']].head())

    # 8. Criação do mapa com Folium (apenas com pontos DENTRO da bacia)
    if not gdf_pontos_filtrados.empty:
        m = folium.Map(location=[-23.98, -46.28], zoom_start=12)

        # Adicionar os polígonos da área do Guarujá ao mapa (opcional, mas bom para contexto)
        folium.GeoJson(
            gdf_area_guaruja.to_json(),
            name='Área de Atuação Guarujá',
            style_function=lambda x: {
                'fillColor': '#0000ff', # Azul
                'color': 'black',
                'weight': 2,
                'fillOpacity': 0.1
            }
        ).add_to(m)

        # Adicionar os pontos ao mapa (apenas os filtrados)
        for idx, row in gdf_pontos_filtrados.iterrows():
            color = 'green' # Cor para pontos validados e dentro da área

            matricula_str = str(row.get('MATRICULA_PADRONIZADA', 'N/A'))
            comunidade_str = str(row.get('COMUNIDADE_PADRONIZADA', 'N/A'))

            folium.CircleMarker(
                location=[row['Latitude_PADRONIZADA'], row['Longitude_PADRONIZADA']],
                radius=5,
                color=color,
                fill=True,
                fill_color=color,
                tooltip=f"Matrícula: {matricula_str}<br>Comunidade: {comunidade_str}"
            ).add_to(m)

        # Adicionar controle de camadas (opcional)
        folium.LayerControl().add_to(m)

        # Salvar o mapa em um arquivo HTML
        m.save("mapa_guaruja_pontos_validos.html")
        print("Mapa interativo com APENAS pontos válidos (sem NaNs e dentro da bacia) criado em 'mapa_guaruja_pontos_validos.html'")
    else:
        print("Nenhum ponto válido encontrado para plotar no mapa após a limpeza de dados e filtragem por bacia.")

    # Próximo passo: Salvar os dados CORRIGIDOS E FILTRADOS
    gdf_pontos_filtrados.to_excel('seus_dados_validos_na_bacia.xlsx', index=False)
    print("\nDados com a coluna 'Dentro_Area_Guaruja' e filtrados (apenas válidos na bacia) salvos em 'seus_dados_validos_na_bacia.xlsx'.")

except FileNotFoundError as e:
    print(f"Erro: Arquivo não encontrado: {e}. Verifique se os arquivos Excel e GeoJSON existem e o caminho está correto.")
except Exception as e:
    print(f"Ocorreu um erro inesperado: {e}")
    import traceback
    traceback.print_exc() # Isso imprimirá o traceback completo para depuração