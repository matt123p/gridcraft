Attribute VB_Name = "Workflow"
Option Explicit
Sub Main()
Dim amount As Double
amount=Range("A1").Value2
If amount>=10 Then amount=amount*0.9
Range("B1").Value2=amount
Debug.Print TypeName(Range("B1").Value2)
End Sub
