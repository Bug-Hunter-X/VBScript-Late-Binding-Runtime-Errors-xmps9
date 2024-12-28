Early Binding:
```vbscript
On Error Resume Next
Dim objExcel As Object
Set objExcel = GetObject(, "Excel.Application")
If Err.Number <> 0 Then
  Err.Clear
  MsgBox "Excel is not running.", vbExclamation
  WScript.Quit
End If
MsgBox objExcel.Version
```
Error Handling:
Always include error handling (On Error Resume Next, Err object) to gracefully handle potential issues.