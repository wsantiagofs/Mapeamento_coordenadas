from geopy.geocoders import Nominatim
import pandas as pd
import time # Importar a biblioteca time para usar sleep
from urllib.parse import quote
import requests

# Inicializa o geocodificador Nominatim com um user_agent e um timeout maior
# Aumentei o timeout para 10 segundos para dar mais tempo à resposta.
geolocator = Nominatim(user_agent="meu_aplicativo_geocodificador_personalizado", timeout=10)

def get_coordinates(address):
    try:
        time.sleep(1.5)
        url = f"https://nominatim.openstreetmap.org/search/{quote(address)}?format=json&addressdetails=1&limit=1"
        headers = {'User-Agent': 'Mozilla/5.0 (compatible; GeocoderBot/1.0)'}
        response = requests.get(url, headers=headers, timeout=10)
        if response.status_code == 200:
            data = response.json()
            if data:
                return float(data[0]['lat']), float(data[0]['lon'])
            else:
                print(f"Nenhum resultado encontrado para o endereço: '{address}'")
                return None, None
        else:
            print(f"Erro HTTP {response.status_code} para o endereço: '{address}'")
            return None, None
    except Exception as e:
        print(f"Erro ao geocodificar o endereço '{address}': {e}")
        return None, None

# Carregue seus dados (substitua 'seu_arquivo.xlsx' pelo caminho correto)
df = pd.read_excel('seus_dados_verificados.xlsx')

# Filtra apenas os 30 primeiros endereços onde 'Dentro_Area_Guaruja' é False
enderecos_false = df[df['Dentro_Area_Guaruja'] == False].head(5).copy()

latitudes = []
longitudes = []
for endereco in enderecos_false['ENDERECO']:
    lat, lon = get_coordinates(endereco)
    latitudes.append(lat)
    longitudes.append(lon)
enderecos_false['Latitude_Nova'] = latitudes
enderecos_false['Longitude_Nova'] = longitudes

print(enderecos_false[['ENDERECO', 'Latitude_Nova', 'Longitude_Nova']])