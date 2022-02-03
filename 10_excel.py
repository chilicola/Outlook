# %%
import os 
import win32com.client as client
from PIL import ImageGrab

workbook_path = os.getcwd() + '\\heatmap.xlsx'
excel = client.Dispatch('Excel.Application')
wb = excel.Workbooks.Open(workbook_path)
sheet = wb.Sheets.Item(1)
# sheet = wb.Sheets[0]
# sheet = wb.Sheets['Sheet1']
excel.visible = 1
copyrange = sheet.Range('A1:M11')
copyrange.CopyPicture(Appearance=1, Format=2)
ImageGrab.grabclipboard().save('paste.png')
excel.Quit()

# %%
image_path = os.getcwd() + '\\paste.png'
html_body = '''
    <div>
        Please review the following report and response with your feedback. <br><br>
    </div>
    <div>
        <img src={}></img>
    </div>
'''

outlook = client.Dispatch('Outlook.Application')
message = outlook.CreateItem(0)
message.To = 'someone@email.com'
message.Subject = 'Please review!'
message.HTMLBody = html_body.format(image_path)
message.Display()
message.Save()
# %%
