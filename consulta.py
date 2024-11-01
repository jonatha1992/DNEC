import os
import requests

api_geminis_url = os.getenv('API_GEMINIS_URL')
api_geminis_key = os.getenv('API_GEMINIS_KEY')
try:
    response = requests.get(f'{api_geminis_url}/api/personal', headers={'api-geminis-key': api_geminis_key})
    response.raise_for_status()
except requests.exceptions.ConnectionError as e:
    print(f"Error de conexión: {e}")
