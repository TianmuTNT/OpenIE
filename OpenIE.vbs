Dim ie
Set ie = CreateObject("InternetExplorer.Application")
ie.Navigate "https://www.bing.com"
ie.Visible = True
Set ie = Nothing