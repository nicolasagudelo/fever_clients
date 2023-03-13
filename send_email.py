import win32com.client

def main(file_path,file_name):
    ol=win32com.client.Dispatch("outlook.application")
    olmailitem=0x0 #size of the new email
    newmail=ol.CreateItem(olmailitem)
    newmail.Subject= 'Fever Clients '+ file_name[:-5]
    newmail.To='nsg@vivaro.com'
    # newmail.CC='maderlyn.machado@vivaro.com;nsg_team@vivaro.com'
    newmail.Body= 'Hello NSG.\n\nFind attached the Fever clients {report}.\n\nPlease remember to check that everything is okay since this file was generated automatically\n\nRegards,\n\nAuto Fever'.format(report = file_name[:-5])
    attach = file_path
    newmail.Attachments.Add(attach)
    # To display the mail before sending it comment the send line.
    newmail.Display() 
    newmail.Save()
    newmail.Send()