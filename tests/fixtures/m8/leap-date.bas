Attribute VB_Name = "Workflow"
Option Explicit
Sub Main()
Range("B1").Value2=CDbl(DateSerial(2024,2,29))
Debug.Print TypeName(Range("B1").Value2)
End Sub
