Attribute VB_Name = "Workflow"
Option Explicit
Sub Main()
Dim a(-1 To 1) As Long,i As Long,n As Long
For i=-1 To 1
a(i)=i+2
n=n+a(i)
Next i
Range("B1").Value2=n
Debug.Print TypeName(Range("B1").Value2)
End Sub
