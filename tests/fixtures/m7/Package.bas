Attribute VB_Name = "Package"
Option Explicit
Public sink As AppSink
Public Sub Main()
Attribute Main.VB_ProcData.VB_Invoke_Func = " \n14"
Application.EnableEvents = False
Dim ws As Object
Set ws = ThisWorkbook.Worksheets("Data")
Range("A1").Value2 = 11
With ws
.Range("A2").Value2 = 22
Range("A3").Value2 = 33
.Range(.Cells(4, 1), .Cells(4, 2)).Value2 = 44
End With
Dim instance As Counter
Set instance = New Counter
instance.Number = 301
ws.Range("A9").Value2 = instance.Read()
ws.Range("A5").Value2 = VarType(CCur(1))
On Error Resume Next
ws.Range("A6").Value2 = 1 / 0
ws.Range("A6").Value2 = Err.Number
End Sub
Public Sub Helper()
Range("A7").Value2 = 77
End Sub
Public Sub Shadow()
Dim Range As Object
Set Range = ThisWorkbook.Worksheets("Data").Range("A8")
Range.Value2 = 88
End Sub
Public Sub Auto_Open()
Range("Z1").Value2 = "must not auto-run"
End Sub
Public Sub Mixed()
Application.EnableEvents = False
On Error Resume Next
With ThisWorkbook.Worksheets("Data")
.Range(Cells(10, 1), Cells(10, 2)).Value2 = 999
.Range("A10").Value2 = Err.Number
End With
End Sub
Public Sub Live()
Application.EnableEvents = False
Application.Range("D1").Value2 = 401
Stop
Range("D2").Value2 = 402
End Sub
Public Sub Connect()
Set sink = New AppSink
sink.Connect
End Sub
