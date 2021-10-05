# %%
import csv
import time
import win32com.client as client

with open('people.csv', newline='') as f:
    reader = csv.reader(f)
    distro = [row for row in reader]

template = '{}, please submit your time as soon as possible'
outlook = client.Dispatch('Outlook.Application')
for name, address in distro:
    message = outlook.CreateItem(0)
    message.To = address
    message.Subject = 'Your time entry is past due!'
    message.Body = template.format(name)
    message.Save()

namespace = outlook.GetNameSpace('MAPI')
drafts = namespace.GetDefaultFolder(16)
messages = list(drafts.Items)
chunks = [messages[x:x+30] for x in range(0, len(messages), 30)]
for chunk in chunks:
    for message in chunk:
        message.Send()
    time.sleep(60)

# %%
