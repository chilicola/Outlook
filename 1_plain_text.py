#%%

import win32com.client as client

outlook = client.Dispatch('Outlook.Application')
message = outlook.CreateItem(0)
message.Display()

message.To = 'test@gmail.com'
message.CC = 'test@gmail.com'
message.BCC = 'test@gmail.com'

message.Subject = 'test subject'
message.Body = 'test body'

message.Save()
message.Send()



# %%
