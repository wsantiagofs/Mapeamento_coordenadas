import pandas as pd
from geopy.geocoders import ArcGIS
from geopy.exc import GeocoderTimedOut, GeocoderUnavailable
import os
import time
import folium
import webbrowser

def get_coordinates(address, geolocator, attempts=3):
    """
    Tenta obter as coordenadas de um endereço, com novas tentativas em caso de falha.
    """
    for attempt in range(attempts):
        try:
            location = geolocator.geocode(address, timeout=10)
            if location:
                return (location.latitude, location.longitude)
        except (GeocoderTimedOut, GeocoderUnavailable):
            print(f" -> Timeout/Serviço indisponível. Tentando novamente ({attempt + 1}/{attempts})...")
            time.sleep(2)
        except Exception as e:
            print(f" -> Ocorreu um erro inesperado: {e}")
            break
    return (None, None)

def create_map(df_results, map_filename='mapa_de_enderecos.html'):
    """
    Cria um mapa HTML com o estilo de satélite híbrido (com nomes), ajusta o zoom para os pontos e o abre automaticamente.
    """
    print(f"\nCriando o mapa em '{map_filename}'...")
    
    # Filtra apenas os endereços que foram encontrados
    df_map = df_results[df_results['Latitude'] != 'Não encontrado'].copy()
    
    # Converte as colunas para numérico para podermos usá-las
    df_map['Latitude'] = pd.to_numeric(df_map['Latitude'])
    df_map['Longitude'] = pd.to_numeric(df_map['Longitude'])
    
    if df_map.empty:
        print(" -> Nenhum endereço com coordenadas válidas encontrado para plotar no mapa.")
        return
        
    # Calcula um centro inicial para o mapa
    map_center = [df_map['Latitude'].mean(), df_map['Longitude'].mean()]
    
    # Cria o objeto do mapa, definindo o satélite HÍBRIDO (com nomes) como padrão
    m = folium.Map(
        location=map_center, 
        tiles='https://mt1.google.com/vt/lyrs=y&x={x}&y={y}&z={z}',
        attr='&copy; Google'
    )
    
    # Cria um grupo de marcadores para poder ajustar os limites
    marker_group = folium.FeatureGroup(name="Endereços Encontrados")
    
    # Adiciona um marcador para cada endereço encontrado
    for idx, row in df_map.iterrows():
        folium.Marker(
            location=[row['Latitude'], row['Longitude']],
            popup=f"<b>{row['Endereço']}</b>"
        ).add_to(marker_group)
    
    # Adiciona o grupo de marcadores ao mapa
    marker_group.add_to(m)
    
    # Ajusta o zoom e o centro do mapa para enquadrar todos os pontos
    m.fit_bounds(marker_group.get_bounds())
        
    # Salva o mapa em um arquivo HTML
    m.save(map_filename)
    print(f" -> Mapa criado com sucesso!")

    # Abre o mapa automaticamente no navegador
    try:
        # Gera o caminho absoluto para o arquivo, que é mais confiável
        full_path = os.path.realpath(map_filename)
        webbrowser.open(f'file://{full_path}')
        print(f" -> O mapa '{map_filename}' foi aberto no seu navegador padrão.")
    except Exception as e:
        print(f" -> Não foi possível abrir o mapa automaticamente: {e}")
        print(f" -> Por favor, abra o arquivo '{map_filename}' manualmente.")


def process_and_save_with_formatting():
    """
    Função principal que lê, processa, salva os resultados e cria um mapa.
    """
    print("Iniciando o processo de busca de coordenadas...")
    
    input_filename = 'Codigo_coordenadas.xlsx'
    output_filename = 'resultado_coordenadas.xlsx'

    if not os.path.exists(input_filename):
        print(f"\nERRO: Arquivo '{input_filename}' não encontrado.")
        print("Por favor, coloque o arquivo na mesma pasta que este script.")
        return

    try:
        print(f"\nLendo o arquivo '{input_filename}' (usando a primeira linha como cabeçalho)...")
        df = pd.read_excel(input_filename, header=0)
        
        enderecos_buscados = []
        latitudes_encontradas = []
        longitudes_encontradas = []
        
        geolocator = ArcGIS(user_agent="geocoding_app_arcgis_formatted")
        
        required_cols = ['Municipio', 'Estado']
        missing_cols = [col for col in required_cols if col not in df.columns]
        if missing_cols:
            print(f"\nERRO: As seguintes colunas não foram encontradas no arquivo de entrada: {', '.join(missing_cols)}")
            print("Por favor, verifique se a primeira linha do seu arquivo Excel contém esses cabeçalhos.")
            return
            
        print("\n--- Buscando coordenadas ---")
        for index, row in df.iterrows():
            address_raw = row.iloc[0]
            
            if pd.isna(address_raw) or not str(address_raw).strip():
                print(f"Linha {index + 2}: Endereço vazio, pulando.")
                enderecos_buscados.append('')
                latitudes_encontradas.append('')
                longitudes_encontradas.append('')
                continue

            municipio = str(row['Municipio']).strip()
            estado = str(row['Estado']).strip()
            
            full_address = f"{str(address_raw)}, {municipio}, {estado}, Brazil"
            
            print(f"Buscando: {full_address}")
            
            latitude, longitude = get_coordinates(full_address, geolocator)
            
            enderecos_buscados.append(full_address)
            if latitude and longitude:
                latitudes_encontradas.append(latitude)
                longitudes_encontradas.append(longitude)
                print(f" -> Coordenadas encontradas: ({latitude}, {longitude})")
            else:
                latitudes_encontradas.append('Não encontrado')
                longitudes_encontradas.append('Não encontrado')
                print(" -> Endereço não encontrado.")
            
            time.sleep(1)

        df['Endereço'] = enderecos_buscados
        df['Latitude'] = latitudes_encontradas
        df['Longitude'] = longitudes_encontradas

        print(f"\nSalvando resultados em '{output_filename}' com formatação...")
        writer = pd.ExcelWriter(output_filename, engine='xlsxwriter')
        
        df.to_excel(writer, sheet_name='Resultados', index=False, header=False, startrow=1)
        
        workbook  = writer.book
        worksheet = writer.sheets['Resultados']
        
        header_format = workbook.add_format({
            'bold': True,
            'border': 1,
            'align': 'center',
            'valign': 'vcenter'
        })
        
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
        
        endereco_col_idx = df.columns.get_loc('Endereço')
        worksheet.set_column(endereco_col_idx, endereco_col_idx, 60)
        worksheet.set_column(endereco_col_idx + 1, endereco_col_idx + 2, 18)

        writer.close()
        
        # --- Geração do Mapa ---
        create_map(df)
        
        print(f"\nProcesso finalizado!")

    except Exception as e:
        print(f"\nOcorreu um erro durante o processamento: {e}")


# --- Ponto de Entrada do Script ---
if __name__ == "__main__":
    # Bibliotecas necessárias:
    # pip install pandas geopy openpyxl xlsxwriter folium
    process_and_save_with_formatting()

