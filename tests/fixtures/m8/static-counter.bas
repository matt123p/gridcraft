Attribute VB_Name = "Workflow"
Option Explicit
Sub Main()
Dim a As Long,b As Long
a=NextId()
b=NextId()
Range("B1").Value2=b
Debug.Print TypeName(Range("B1").Value2)
End Sub
Function NextId() As Long
Static n As Long
n=n+1
NextId=n
End Function
