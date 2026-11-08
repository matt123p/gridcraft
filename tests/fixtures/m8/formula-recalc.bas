Attribute VB_Name = "Workflow"
Option Explicit
Sub Main()
Range("B1").Formula="=SUM(A1:A3)"
Debug.Print TypeName(Range("B1").Value2)
End Sub
