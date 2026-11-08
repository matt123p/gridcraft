Attribute VB_Name = "Workflow"
Option Explicit
Sub Main()
Dim v As Variant
v=Null
Range("B1").Value2=IsNull(v)
Debug.Print TypeName(Range("B1").Value2)
End Sub
