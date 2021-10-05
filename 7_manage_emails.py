# %%
from os import name
import win32com.client as client 
outlook = client.Dispatch('Outlook.Application')
namespace = outlook.GetNameSpace('MAPI')
drafts = namespace.GetDefaultFolder(16)
drafts.Items.Count

message = outlook.createitem(0)
message.SenderName
message.SenderEmailAddress
message.Subject = 'testing subject'
message.Body = 'Hey, there.'
# message.Display()
message.save()
# message.delete()

message = drafts.Items[0]
message = drafts.Items.Item(1)

for _ in range(10):
    message.copy()

# %%
