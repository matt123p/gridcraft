Attribute VB_Name = "BookCode"
Attribute VB_PredeclaredId = True
Private Sub Workbook_Open()
Application.EnableEvents = False
Range("C1").Value2 = 201
Me.Worksheets("Data").Range("C2").Value2 = 202
End Sub
