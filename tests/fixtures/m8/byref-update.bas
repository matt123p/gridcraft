Attribute VB_Name = "Workflow"
Option Explicit
Sub Main()
Dim n As Long
n=7
DoubleIt n
Range("B1").Value2=n
Debug.Print TypeName(Range("B1").Value2)
End Sub
Sub DoubleIt(ByRef n As Long)
n=n*2
End Sub
