Attribute VB_Name = "AppSink"
Private WithEvents app As Excel.Application
Public Sub Connect()
Set app = Application
End Sub
Private Sub app_WorkbookBeforeSave(ByVal Wb As Excel.Workbook, ByVal SaveAsUI As Boolean, Cancel As Boolean)
Wb.Worksheets("Data").Range("E1").Value2 = 501
Cancel = True
End Sub
