import win32com.client

# créer une instance de l'application Outlook
outlook = win32com.client.Dispatch("Outlook.Application")

# accéder à la boîte de réception
inbox = outlook.GetNamespace("MAPI").GetDefaultFolder(6)

# parcourir les messages de la boîte de réception
for message in inbox.Items:
    print(message.Subject)
