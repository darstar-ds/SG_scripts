from elevenlabs.api import Models
from elevenlabs import set_api_key
import os


ELEVENLABS_APIKEY = os.environ.get("API_DS_ELEVENLABS")

set_api_key(str(ELEVENLABS_APIKEY))
# print(type(ELEVENLABS_APIKEY))
models = Models.from_api()
# print(models[0])
print(f"Liczba modeli: {len(models)}")
print(models)