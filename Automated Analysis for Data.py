while True:
    
    try:
        
        import datetime
        import os
        import win32com.client


#         path = os.path.expanduser("~/Desktop/Attachments")
        today = datetime.date.today()

        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        inbox = outlook.GetDefaultFolder(6) 
        messages = inbox.Items


        def saveattachments(subject):
            for message in messages:
                if message.Subject == subject and message.UnRead :
                    # body_content = message.body
                    attachments = message.Attachments
                    attachment = attachments.Item(1)
                    for attachment in message.Attachments:

                        attachment.SaveAsFile(r"C:\Users\MISHS\OneDrive -\Documents\Job_.htm")
                        global path
                        path= r"C:\Users\MISHS\OneDrive -\Documents\Job_.htm"
                        if message.Subject == subject and message.UnRead:
                            message.UnRead = False
                        break


        saveattachments('Job REPORT, Step 1')    


        import pandas as pd

        df = pd.read_html(path, flavor='bs4')
        consolidated_df= pd.DataFrame()
        consolidated_df= consolidated_df.append(df[0])
        conso_df= pd.DataFrame()
        cols= conso_df.iloc[0,:]
        conso_df.columns= cols
        conso_df.drop(0,axis=0, inplace= True)
        conso_df.to_excel("C:\\Users\\MISHS\\OneDrive -\\Documents\\Job___.xlsx")

        import win32com.client as win32
        import pandas as pd
        from pathlib import Path
        from datetime import date


        to_email = 'shubham.mishra@*****.com'

        df = pd.read_excel(r"C:\\Users\\MISHS\\OneDrive -\\Documents\\Job___.xlsx")


        outlook = win32.gencache.EnsureDispatch('Outlook.Application')
        new_mail = outlook.CreateItem(0)


        new_mail.Subject = "{:%m/%d} Job MLL Report Update".format(date.today())


        new_mail.To = to_email

        attachment1 = r"C:\\Users\\MISHS\\OneDrive -\\Documents\\Job___.xlsx"

        new_mail.Attachments.Add(Source=str(attachment1))

        new_mail.Display(True)
        
    except :
        print("No new unread mail as of now")
        continue
