from msg_parser import MsOxMessage

msg_obj = MsOxMessage("./SAP Headsup/emails/MARKETING - 7031-1  Heads-up.msg")

# is_message = msg_obj.is_valid_msg_file()
# print(is_message)
json_string = msg_obj.get_message_as_json()

print(type(json_string))
# print(json_string)

msg_properties_dict = msg_obj.get_properties()
