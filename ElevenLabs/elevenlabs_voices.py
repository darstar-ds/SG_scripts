from elevenlabs.api import Voices
from elevenlabs import set_api_key
import os


ELEVENLABS_APIKEY = os.environ.get("API_DS_ELEVENLABS")

set_api_key(str(ELEVENLABS_APIKEY))

voices = Voices.from_api()
# print(type(voices))
print(voices[0].name)
# print(voices)
# print(voices[0][1]) #TypeError: 'Voice' object is not subscriptable

for voice in voices:
    print(f"Voice name: {voice.name}")
