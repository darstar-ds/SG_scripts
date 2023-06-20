# Note: you need to be using OpenAI Python v0.27.0 for the code below to work
import openai
import os

API_OPENAI_KEY = os.environ.get("API_DS_OPENAI")

openai.api_key = API_OPENAI_KEY

audio_file = open("./Whisper/_audio_file/DS_derush_po15.mp3", "rb")
transcript = openai.Audio.transcribe("whisper-1", audio_file, response_format="srt", language="pl")

# transcript.write("./Whisper/_audio_file/subtitle.srt")

print(type(transcript))
print(transcript)

# with open(transcript, "w") as text_file:
#     text_file.write("./Whisper/_audio_file/subtitle.srt")

print(transcript, file=open("./Whisper/_audio_file/subtitle.srt", "w"))