from appscript import app, k
from mactypes import Alias

# Mac OS / Microsoft Outlook app support only. Can look into Office365-REST-Python-Client later
#
def compose_outlook_email(subject, recipient_name, recipient_email, body, attachment=''):
    outlook = app('Microsoft Outlook')
    msg = outlook.make(new=k.outgoing_message, with_properties={k.subject: subject, k.plain_text_content: body})
    msg.make(new=k.recipient, with_properties={k.email_address: {k.name: recipient_name, k.address: recipient_email}})
    if attachment:
        msg.make(new=k.attachment, with_properties={k.file: Alias(attachment)})
    msg.open()
    msg.activate()
