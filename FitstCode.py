import win32com.client as win32

xl = win32.DispatchEx("Excel.Application")
xl.Visible = True
wb = xl.Workbooks.Add()
model = wb.VBProject.VBComponents.Add(1)

VBACode = '''
Sub Python()
    MsgBox "Hello Excel VBA"
End Sub
'''

model.CodeModule.AddFromString(VBACode)
xl.Run("Python")
