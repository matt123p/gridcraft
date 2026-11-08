Attribute VB_Name = "Workflow"
Option Explicit
Sub Main()
Dim total As Currency
total=CCur(6.65)*3
Range("B1").Value2=total
Debug.Print TypeName(Range("B1").Value2)
End Sub
