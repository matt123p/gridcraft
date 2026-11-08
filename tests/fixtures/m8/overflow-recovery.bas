Attribute VB_Name = "Workflow"
Option Explicit
Sub Main()
On Error Resume Next
Dim n As Integer
n=40000
Range("B1").Value2=Err.Number
Debug.Print TypeName(Range("B1").Value2)
End Sub
