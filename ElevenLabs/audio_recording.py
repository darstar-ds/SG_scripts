from elevenlabs import generate, play, save
import os

ELEVENLABS_APIKEY = os.environ.get("API_DS_ELEVENLABS")
RECORDING_FOLDER = "d:\\Moje dokumenty\\SG_scripts_data\\SAP\\recording\\audio\\"
PROMPTS_FOLDER = "d:\\Moje dokumenty\\SG_scripts_data\\SAP\\recording\\"
RECORDING_VOICE = "Bella"
RECORDING_PROMPTS = PROMPTS_FOLDER + "Language_Prompts_v5.txt"

def record_audio(prompt):
    audio = generate(
            text=prompt,
            api_key= ELEVENLABS_APIKEY,
            voice= RECORDING_VOICE,
            model='eleven_multilingual_v1'
            )
    return audio

def save_recording(audio_rec, file_name):
    save(
    audio= audio_rec,
    filename= RECORDING_FOLDER + file_name
    )

with open(RECORDING_PROMPTS) as rec_prompts:
    lines = [line.rstrip() for line in rec_prompts]
    # print(lines)
    audio = record_audio(lines[0])
    save_recording(audio, "nazwa0.wav")
# play(audio)

