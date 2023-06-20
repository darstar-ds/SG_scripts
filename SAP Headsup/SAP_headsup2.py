import extract_msg
import pandas as pd

def table2dictlist(quasitable):
    # Split the text into lines
    lines = quasitable.splitlines()
    # print(lines)
    # Remove the header row
    header = lines[2]
    print(f"Headers: {header}")
    lines = lines[3:-2]

    # Split each line into columns and create a dictionary for each row
    data_list = []
    for line in lines:
        cols = line.split('\t')
        data_dict = {}
        for i, col in enumerate(cols):
            header_name = header.split('\t')[i].strip()
            data_dict[header_name] = col.strip()
        data_list.append(data_dict)

    return data_list


# open message
msg = extract_msg.Message("./SAP Headsup/emails/MARKETING - 7031-1  Heads-up.msg")
# print sender name
print('Sender: {}'.format(msg.sender))
# print date
print('Sent On: {}'.format(msg.date))
# print subject
print('Subject: {}'.format(msg.subject))
# print body
print(type(msg.body))

with open("./SAP Headsup/timeline_info.txt", "w") as file:
    email_body = msg.body
    timeline_info = email_body.split("Timeline information:")[1].split("Volume information:")[0].split("Back to top")[0]
    # print(timeline_info)
    timeline_dict_list = table2dictlist(timeline_info)
    df_timeline = pd.DataFrame(timeline_dict_list)
    file.write(timeline_info)

with open("./SAP Headsup/volume_info.txt", "w") as file:
    email_body = msg.body
    volume_info = email_body.split("Volume information:")[1].split("Back to top")[0]
    # print(volume_info)
    volume_dict_list = table2dictlist(volume_info)
    df_volume = pd.DataFrame(volume_dict_list)
    file.write(volume_info)

# print('Body: {}'.format(msg.body))
print(df_timeline)
print(df_volume)

# Merge the two dataframes on the 'Language' and 'Step' columns
merged_df = pd.merge(df_timeline, df_volume, on=['Language', 'Step'])
print(merged_df)
print(merged_df.info)
merged_df.describe()
