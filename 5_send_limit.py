# %%

import win32com.client as client

distro = ['test@gmail.com' for _ in range(1567)]

chunks = []
for x in range(0, len(distro), 500):
    chunk = distro[x : x+500]
    chunks.append(chunk)

outlook = client.Dispatch('Outlook.Application')
for recipients in chunks:
    message = outlook.CreateItem(0)
    message.To = ';'.join(recipients)
    message.Subject = 'Missing time alert!'
    message.Body = 'Please submit your time as soon as possible!'
    message.Save()


# %%
