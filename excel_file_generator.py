import win32com.client as win32

# Create an instance of word
WordApp = win32.gencache.EnsureDispatch('Excel.Application')
WordApp
# Make the app visible
WordApp.Visible = True
WordDoc = WordApp.Documents.Add()

WordDoc = win32.gencache.EnsureDispatch(WordApp.Documents(1))

help(WordDoc)
