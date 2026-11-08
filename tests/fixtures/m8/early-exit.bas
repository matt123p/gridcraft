Attribute VB_Name = "Workflow"
Option Explicit
Sub Main()
Dim i As Long
For i=1 To 10
If i Mod 3=0 Then Exit For
Next i
Range("B1").Value2=i
Debug.Print TypeName(Range("B1").Value2)
End Sub
