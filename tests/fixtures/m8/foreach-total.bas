Attribute VB_Name = "Workflow"
Option Explicit
Sub Main()
Dim v As Variant,n As Long,item As Variant
v=Array(1,2,3,4)
For Each item In v
n=n+item
Next item
Range("B1").Value2=n
Debug.Print TypeName(Range("B1").Value2)
End Sub
