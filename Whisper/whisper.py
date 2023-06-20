# Note: you need to be using OpenAI Python v0.27.0 for the code below to work
import openai

openai.api_key = "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"

audio_file = open("./Whisper/_audio_file/DS_derush_po15.mp3", "rb")
transcript = openai.Audio.transcribe("whisper-1", audio_file, response_format="srt", language="pl")

# transcript.write("./Whisper/_audio_file/subtitle.srt")

print(type(transcript))
print(transcript)

# with open(transcript, "w") as text_file:
#     text_file.write("./Whisper/_audio_file/subtitle.srt")

print(transcript, file=open("./Whisper/_audio_file/subtitle.srt", "w"))