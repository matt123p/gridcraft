Attribute VB_Name = "Workflow"
Option Explicit
Sub Main()
Dim n As Long
n=Range("A2").Value2
If n>=10 And n<=20 Then
Range("B1").Value2="medium"
Else
Range("B1").Value2="other"
End If
Debug.Print TypeName(Range("B1").Value2)
End Sub
