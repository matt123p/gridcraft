Attribute VB_Name = "SheetCode"
Attribute VB_PredeclaredId = True
Private Sub Worksheet_Change(ByVal Target As Range)
Application.EnableEvents = False
Range("B1").Value2 = 101
Me.Cells(2, 2).Value2 = 102
Range("B6").Value2 = Rows(6).Row
Range("B7").Value2 = Columns(3).Column
Helper
SheetCode.Range("B3").Value2 = 103
With ThisWorkbook.Worksheets("Other")
.Range("B4").Value2 = 104
Range("B5").Value2 = 105
End With
End Sub
