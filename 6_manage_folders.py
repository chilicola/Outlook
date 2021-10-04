# %%
import win32com.client as client
outlook = client.Dispatch('Outlook.Application')

namespace = outlook.GetNameSpace('MAPI')
drafts = namespace.GetDefaultFolder(16)
inbox = namespace.GetDefaultFolder(6)

testfolder = inbox.Folders['test1']
testfolder.Name
testfolder = inbox.Folders[0]
testfolder.Name
testfolder = inbox.Folders.Item(1)
testfolder.Name

for folder in inbox.Folders:
    print(folder.Name)

inbox.Folders.Count

# inbox.Folders.Add('test4')

inbox.Name
inbox.Description
inbox.Folders['test1'].Description = 'description in folder test1'
inbox.Folders['test1'].Description

inbox.FolderPath

parent = testfolder.Parent
parent.Name

inbox.Items[0].Subject

newfolder = inbox.Folders.Add('MyNewFolder')
# newfolder.MoveTo(testfolder)
newfolder.Delete()
# %%
