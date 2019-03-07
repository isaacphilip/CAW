Option Explicit
Dim objExcel, objDOCMAN, strExcelPath, strDOCMAN, objSheet, objSheet2, rows, col, row, words, a, x
Dim stresetup, objesetup, objSheete
Dim strHome, objHome, rowp, strdocit, rowd, objdocit, objShell


Set objShell = CreateObject("Wscript.Shell")
objShell.Run("taskkill /im EXCEL.exe"), 1, TRUE
objShell.Run("taskkill /im EXCEL.exe"), 1, TRUE
objShell.Run("taskkill /im EXCEL.exe"), 1, TRUE
objShell.Run("taskkill /im EXCEL.exe"), 1, TRUE
objShell.Run("taskkill /im iexplore.exe"), 1, TRUE
WScript.Sleep 2000

strExcelPath = "C:\BMS Automation\eSetup log\eSetupData.xlsx"
stresetup = "C:\BMS Automation\eSetup Approve\data.xlsx"
strDOCMAN = "C:\BMS Automation\DOCMAN\data.xlsx"


Set objExcel = CreateObject("Excel.Application")
objExcel.WorkBooks.Open strExcelPath
Set objSheet = objExcel.ActiveWorkbook.Worksheets(1)


DIM IE
DIM urls
Set IE = CreateObject("InternetExplorer.Application")

rows = 1
words = 0
col = 1
row = 1
rowp = 1
rowd = 1


Set objesetup = CreateObject("Excel.Application")
objesetup.WorkBooks.Open stresetup
Set objSheete = objesetup.ActiveWorkbook.Worksheets(1)

Do Until objExcel.Cells(rows,1).Value =  ""
 If (Left(objExcel.Cells(rows,1).Value, 30) =  "DOCMAN [Add] [Accenture (ACN)]") Then
  objSheete.Cells(rows, 1).Value = Left(objExcel.Cells(rows,1).Value, 39)
  objSheete.Cells(rows, 2).Value = objSheet.Cells(rows, 2).Value
  objesetup.ActiveWorkbook.Save
  IE.Visible = 1
  Set objDOCMAN = CreateObject("Excel.Application")
  objDOCMAN.WorkBooks.Open strDOCMAN
  Set objSheet2 = objDOCMAN.ActiveWorkbook.Worksheets(1)
  IE.Navigate objSheet.Cells(rows, 2).Value
  While IE.ReadyState <> 4
    WScript.Sleep 1000
  Wend
  set urls = ie.document.all.tags("a")
  col = 1
  WScript.Sleep 1000
  Do Until words > 150 OR IE.document.getElementsByTagName("font").Item(words).InnerText =  "-"
   If col = 4 Then
    objSheet2.Cells(rowp, col).Value =Mid(urls(2).innerHTML,7,9)
    objDOCMAN.ActiveWorkbook.Save
    col = col+1
   Else
    If(words < 4 OR words > 10) Then
     objSheet2.Cells(rowp, col).Value = Trim(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(IE.document.getElementsByTagName("font").Item(words).InnerText,"<"," "),">"," "),"(UID="," "),")", " "),"Select Cabinet:"," "),"Action:"," "),"Select application:"," "),"Select Role:"," "))
     objDOCMAN.ActiveWorkbook.Save
     col = col+1
    End If
    words= words +1
   End If
  Loop
  rowp=rowp+1
  objDOCMAN.ActiveWorkbook.Save
  objDOCMAN.ActiveWorkbook.Close
  objDOCMAN.Application.Quit
  objDOCMAN.Quit
 End If
 
 
 rows = rows +1
 words= 0
Loop
objesetup.ActiveWorkbook.Save
objesetup.ActiveWorkbook.Close
objesetup.Application.Quit
objesetup.Quit
objExcel.ActiveWorkbook.Close
objExcel.Application.Quit
objExcel.Quit
objShell.Run("taskkill /im EXCEL.exe"), 1, TRUE
Msgbox "completed"
