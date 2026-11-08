Attribute VB_Name = "Workflow"
Option Explicit
Sub Main()
Dim i As Long,n As Double,v As Double
For i=1 To 3
v=Range("A" & CStr(i)).Value2
If v>0 Then n=n+v
Next i
Range("B1").Value2=n
Debug.Print TypeName(Range("B1").Value2)
End Sub
