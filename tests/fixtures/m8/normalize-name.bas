Attribute VB_Name = "Workflow"
Option Explicit
Sub Main()
Range("B1").Value2=UCase$(Trim$(Range("A4").Value2))
Debug.Print TypeName(Range("B1").Value2)
End Sub
