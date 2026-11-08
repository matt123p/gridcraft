Attribute VB_Name = "Workflow"
Option Explicit
Sub Main()
Range("B1").Value2=Len(Range("A5").Value2)
Debug.Print TypeName(Range("B1").Value2)
End Sub
