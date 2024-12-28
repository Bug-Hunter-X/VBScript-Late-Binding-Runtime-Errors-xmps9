Late Binding: VBScript's late binding can lead to runtime errors if an object or method doesn't exist.  This is harder to catch during development.
```vbscript
Dim objExcel
Set objExcel = CreateObject("Excel.Application")
'This will fail if Excel isn't installed
MsgBox objExcel.Version
```