Option Explicit
Dim objExcel, objExcel1, objShell, objWorkbook, objPDHQ, strExcelPath
Dim strPDHQ, objSheet, objSheet2, rows, col, row, words, a, x
Dim objbook, IE, intRow, struid, stresetup, strsite, oShell
Dim FSO, f, name, objTextFile

Set FSO =CreateObject("scripting.FileSystemObject")

For Each f in FSO.GetFolder("C:\BMS Automation\log").Files
  name = LCase(f.Name)
  If FSO.GetExtensionName(name) = "txt" Then
    Set objTextFile = FSO.OpenTextFile ("C:\BMS Automation\log\"& f.Name, 8, True)
  End If
Next
a = Now()
objTextFile.WriteLine(a)
objTextFile.WriteLine("Added eSetup number :"& stresetup)

objTextFile.Close