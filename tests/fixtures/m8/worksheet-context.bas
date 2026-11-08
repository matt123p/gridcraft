Attribute VB_Name = "Workflow"
Option Explicit
Sub Main()
Dim ws As Excel.Worksheet
Set ws=ThisWorkbook.Worksheets("Data")
ws.Range("B1").Value2=ws.Range("A1").Value2
Debug.Print TypeName(Range("B1").Value2)
End Sub
