# %%
import win32com.client as client
outlook = client.Dispatch('Outlook.Application')
namespace = outlook.GetNameSpace('MAPI')
# account = namespace.Folders['tingtzuhao@gmail.com']
# inbox = namespace.GetDefaultFolder(6)
account = namespace.Folders['kevin@webmainland.com']
inbox = account.Folders['Inbox']

yt_emails = [message for message in inbox.Items if message.SenderEmailAddress.endswith('youtube.com')]
for message in yt_emails:
    print(message)

yt_folder = inbox.Folders.Add('YTEmails')
for message in yt_emails:
    message.Move(yt_folder)

junk_messages = [message for message in inbox.Items if 'deal' in message.Body.lower()]
for message in junk_messages:
    print(message.SenderEmailAddress)

for message in junk_messages:
    print(message.Subject)

junk_stuff = inbox.Folders.Add('JunkStuff')
for message in junk_messages:
    message.Move(junk_stuff)

def mail_body_search(term, folder):
    '''Recursively search all folders for email containing search term'''
    relevant_messages = [(message, message.Parent.Name) for message in folder.Items if term in message.Body.lower()]

    # check for subfolders (base case)
    subfolder_count = folder.Folders.Count

    # search all subfolders
    if subfolder_count > 0:
        for subfolder in folder.Folders:
            relevant_messages.extend(mail_body_search(term, subfolder))
    return relevant_messages

# search for python in my account folder
results = mail_body_search('python', account)

# %%
import csv
with open('search_results.csv', 'w', newline='', encoding='utf-8') as f:
    writer = csv.writer(f)
    writer.writerow(['ParentFolder', 'Subject'])
    for message, parent in results:
        writer.writerow([parent, message.Subject])