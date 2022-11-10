import win32com.client as win32

olApp = win32com.Dispatch("Outlook.Application")
olNS = olApp.getNameSpace("MAPI")

mailItem = olApp.CreateItem(0)
mailItem.Subject = "Hello 123"
mailItem.BodyFormat = 1
mailItem.Body = "Hello There"
mailItem.To = "stianroger@gmail.com"
mailItem._oleobj_.Invoke(*(64209, 0, 8, 0, olNS.Accounts.Item("Lars27055@gmail.com")))

mailItem.BodyFormat = 2
mailItem.HTMLBody = "<HTML Markup>"


