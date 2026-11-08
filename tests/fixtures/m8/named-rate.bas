Attribute VB_Name = "Workflow"
Option Explicit
Sub Main()
Dim n As Object
Set n=ThisWorkbook.Names.Add(Name:="Rate", RefersTo:="=0.25")
Range("B1").Value2=Evaluate("Rate*8")
Debug.Print TypeName(Range("B1").Value2)
End Sub
