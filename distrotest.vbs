Option Explicit
Dim objExcel, objPDHQ, strExcelPath, strPDHQ, objSheet, objSheet2, rows, col, row, words, a, x
Dim stresetup, objesetup, objSheete, objShell, strCurDir, rowp
Dim fso, objconfig

Set objShell = CreateObject("Wscript.Shell")
'objShell.Run("taskkill /im EXCEL.exe"), 1, TRUE
'objShell.Run("taskkill /im EXCEL.exe"), 1, TRUE
'objShell.Run("taskkill /im EXCEL.exe"), 1, TRUE
'objShell.Run("taskkill /im EXCEL.exe"), 1, TRUE
'objShell.Run("taskkill /im iexplore.exe"), 1, TRUE
'WScript.Sleep 2000

Set fso = CreateObject("Scripting.FileSystemObject")

strCurDir = objShell.CurrentDirectory 

Set objconfig = CreateObject("Excel.Application")
objconfig.WorkBooks.Open ("C:\BMS Automation\eSetup log\config.xlsx")
Set objSheet2 = objconfig.ActiveWorkbook.Worksheets(1)

Set objExcel = CreateObject("Excel.Application")
objExcel.WorkBooks.Open objconfig.Cells(1,2).Value
Set objSheet = objExcel.ActiveWorkbook.Worksheets(1)
stresetup = objconfig.Cells(2,2).Value

'Wscript.Echo stresetup

DIM IE
DIM urls
Set IE = CreateObject("InternetExplorer.Application")

rows = 1
words = 0
col = 1
row = 1
rowp= 1

Set objesetup = CreateObject("Excel.Application")
objesetup.WorkBooks.Open stresetup
Set objSheete = objesetup.ActiveWorkbook.Worksheets(1)

Do Until objExcel.Cells(rows,1).Value =  ""
 Do Until objconfig.Cells(row,1).Value =  ""
  If (objExcel.Cells(rows,1).Value =  objconfig.Cells(row,1).Value) Then
   objSheete.Cells(rows, 1).Value = objExcel.Cells(rows,1).Value
   objSheete.Cells(rows, 2).Value = objSheet.Cells(rows, 2).Value
   objesetup.ActiveWorkbook.Save
   IE.Visible = 1
   Set objPDHQ = CreateObject("Excel.Application")
   objPDHQ.WorkBooks.Open objconfig.Cells(row,2).Value
   Set objSheet2 = objPDHQ.ActiveWorkbook.Worksheets(1)
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
     objPDHQ.ActiveWorkbook.Save
	 Wscript.Echo objSheet2.Cells(rowp, col).Value
     col = col+1
    Else
     If(words < 4 OR words > 10) Then
      objSheet2.Cells(rowp, col).Value = Trim(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(IE.document.getElementsByTagName("font").Item(words).InnerText,"<"," "),">"," "),"(UID="," "),")", " "),"option:"," "),"Action:"," "),"Select application:"," "),"Select the appropriate role:"," "),"= (", " "),"Roles:"," "),"Environment:"," "),"Site:"," "),"="," "),"Digital Certificate Type:"," "),"Select a Project: "," "),"Select Group/Role"," "),"Select role/option:"," "),"MS Project Pro Needed?:"," "),"Role:"," "),"Specify Role (EDC Test Only :"," "),"Specify Role (EDC :"," "),"Alliance Groups: (hold CTRL to multi-select"," "),"Specify"," "),"country:"," "),"Division:"," "),"Country:"," "),"Select"," "),"Application(s :"," "),"Type:"," "))
      objPDHQ.ActiveWorkbook.Save
	  Wscript.Echo objSheet2.Cells(rowp, col).Value
      col = col+1
     End If
     words= words +1
    End If
   Loop
   rowp=rowp+1
   objPDHQ.ActiveWorkbook.Save
   objPDHQ.ActiveWorkbook.Close
   objPDHQ.Application.Quit
   objPDHQ.Quit
   
   'rows = rows + 1
  End If
  row = row +1
  words = 0
 Loop
 rows = rows + 1
 row = 1
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