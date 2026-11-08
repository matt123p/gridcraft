Attribute VB_Name = "Workflow"
Option Explicit
Sub Main()
Range("B1").Value2 = Range("A1").Value2 * Range("A2").Value2
Debug.Print TypeName(Range("B1").Value2)
End Sub
