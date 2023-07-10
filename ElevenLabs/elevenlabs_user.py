from elevenlabs.api import User
from elevenlabs import set_api_key
import os


ELEVENLABS_APIKEY = os.environ.get("API_DS_ELEVENLABS")

set_api_key(str(ELEVENLABS_APIKEY))
# print(type(ELEVENLABS_APIKEY))
user = User.from_api()
print(user)