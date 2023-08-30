from transformers import T5Tokenizer, T5ForConditionalGeneration

model = T5ForConditionalGeneration.from_pretrained("Voicelab/vlt5-base-keywords")
tokenizer = T5Tokenizer.from_pretrained("Voicelab/vlt5-base-keywords")

task_prefix = "Keywords: "
inputs = [
    'BY USING THE SPOTIFY SERVICE, YOU AFFIRM THAT YOU ARE 18 YEARS OR OLDER TO ENTER INTO THESE TERMS, OR, IF YOU ARE NOT, THAT YOU ARE 13 YEARS OR OLDER AND HAVE OBTAINED PARENTAL OR GUARDIAN CONSENT TO ENTER INTO THESE TERMS.',
    'Additionally, in order to use the Spotify Service and access any Content, you represent that: you reside in the United States, and any registration and account information that you submit to Spotify is true, accurate, and complete, and you agree to keep it that way at all times.',
    'We provide numerous Spotify Service options.',
    'Certain Spotify Service options are provided free of charge, while other options require payment before they can be accessed ("Paid Subscriptions").',
    'We may also offer special promotional plans, memberships, or services, including offerings of third-party products and services.',
    'We are not responsible for the products and services provided by such third parties.',
    'The Spotify Service may be integrated with, or may otherwise interact with, third-party applications, websites, and services ("Third-Party Applications") and third-party personal computers, mobile handsets, tablets, wearable devices, speakers, and other devices ("Devices"). ',
    'Your use of such Third-Party Applications and Devices may be subject to additional terms, conditions, and policies provided to you by the applicable third party. ',
    'Spotify does not guarantee that Third-Party Applications and Devices will be compatible with the Spotify Service.',
]

for sample in inputs:
    input_sequences = [task_prefix + sample]
    input_ids = tokenizer(
        input_sequences, return_tensors="pt", truncation=True
    ).input_ids
    output = model.generate(input_ids, max_new_tokens=20, no_repeat_ngram_size=3, num_beams=4)
    predicted = tokenizer.decode(output[0], skip_special_tokens=True)
    print(sample, "\n --->", predicted)
