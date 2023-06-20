from pydub import AudioSegment

audio_file = AudioSegment.from_mp3("./Whisper/_audio_file/derush.mp3")

# PyDub handles time in milliseconds
fifteen_minutes = 14 * 60 * 1000 + 51 * 1000

first_15_minutes = audio_file[:fifteen_minutes]
after_15_minutes = audio_file[fifteen_minutes:]
first_minute = audio_file[:60000]

first_15_minutes.export("./Whisper/_audio_file/DS_derush_do15.mp3", format="mp3")
after_15_minutes.export("./Whisper/_audio_file/DS_derush_po15.mp3", format="mp3")
first_minute.export("./Whisper/_audio_file/DS_derush_first_minute.mp3", format="mp3")