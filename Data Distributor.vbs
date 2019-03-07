Option Explicit
Dim objExcel, objPDHQ, strExcelPath, strPDHQ, objSheet, objSheet2, rows, col, row, words, a, x
Dim stresetup, objesetup, objSheete
Dim strHome, objHome, rowp, strdocit, rowd, objdocit, objShell
Dim strdevlim, objdevlim, rowde, strdigi, rowdi, objdigi
Dim strtab, objtab, rowtab
Dim strnbr, objnbr, rownbr
Dim strdep, objdep, rowdep
Dim strprism, objprism, rowprism
Dim strppm, objppm, rowppm
Dim strtao, objtao, rowtao
Dim strspot, objspot, rowspot
Dim strplex, objplex, rowplex
Dim strjr, objjr, rowjr
Dim strdim, objdim, rowdim
Dim strcisco, objcisco, rowcisco
Dim strgts, objgts, rowgts
Dim strtcana, objtcana, rowtcana
Dim strautodsk, objautodsk, rowautodsk
Dim strdeltav, objdeltav, rowdeltav

Set objShell = CreateObject("Wscript.Shell")
objShell.Run("taskkill /im EXCEL.exe"), 1, TRUE
objShell.Run("taskkill /im EXCEL.exe"), 1, TRUE
objShell.Run("taskkill /im EXCEL.exe"), 1, TRUE
objShell.Run("taskkill /im EXCEL.exe"), 1, TRUE
objShell.Run("taskkill /im iexplore.exe"), 1, TRUE
WScript.Sleep 2000

strExcelPath = "C:\BMS Automation\eSetup log\eSetupData.xlsx"
stresetup = "C:\BMS Automation\eSetup Approve\data.xlsx"
strPDHQ = "C:\BMS Automation\PDHQ\data.xlsx"
strHome = "C:\BMS Automation\Home Directory Automation\data.xlsx"
strdocit = "C:\BMS Automation\DOCIT\data.xlsx"
strdevlim = "C:\BMS Automation\Devlims\data.xlsx"
strdigi = "C:\BMS Automation\Digital Certificate\data.xlsx"
strtab = "C:\BMS Automation\Tableau\data.xlsx"
strnbr = "C:\BMS Automation\NBR RTReports\data.xlsx"
strdep = "C:\BMS Automation\Departures\data.xlsx"
strprism = "C:\BMS Automation\PRISM CARA\data.xlsx"
strppm = "C:\BMS Automation\PPM\data.xlsx"
strtao = "C:\BMS Automation\TAO\data.xlsx"
strspot = "C:\BMS Automation\Spotfire\data.xlsx"
strplex = "C:\BMS Automation\Plexus\data.xlsx"
strjr = "C:\BMS Automation\Jreview\data.xlsx"
strdim = "C:\BMS Automation\DIMS\data.xlsx"
strcisco = "C:\BMS Automation\CISCO\data.xlsx"
strgts = "C:\BMS Automation\GTS\data.xlsx"
strtcana = "C:\BMS Automation\TCANA\data.xlsx"
strautodsk = "C:\BMS Automation\Autodesk\data.xlsx"
strdeltav = "C:\BMS Automation\DeltaV\data.xlsx"



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
rowdi = 1
rowde = 1
rowtab = 1
rownbr = 1
rowdep = 1
rowprism = 1
rowppm = 1
rowtao = 1
rowspot = 1
rowplex = 1
rowjr = 1
rowdim = 1
rowcisco = 1
rowgts = 1
rowtcana = 1
rowautodsk = 1
rowdeltav = 1

Set objesetup = CreateObject("Excel.Application")
objesetup.WorkBooks.Open stresetup
Set objSheete = objesetup.ActiveWorkbook.Worksheets(1)

Do Until objExcel.Cells(rows,1).Value =  ""
 If (Left(objExcel.Cells(rows,1).Value, 39) =  "PROCEDURAL DOCUMENT HEADQUARTERS (PDHQ)") Then
  objSheete.Cells(rows, 1).Value = Left(objExcel.Cells(rows,1).Value, 39)
  objSheete.Cells(rows, 2).Value = objSheet.Cells(rows, 2).Value
  objesetup.ActiveWorkbook.Save
  IE.Visible = 1
  Set objPDHQ = CreateObject("Excel.Application")
  objPDHQ.WorkBooks.Open strPDHQ
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
    col = col+1
   Else
    If(words < 4 OR words > 10) Then
     objSheet2.Cells(rowp, col).Value = Trim(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(IE.document.getElementsByTagName("font").Item(words).InnerText,"<"," "),">"," "),"(UID="," "),")", " "),"option:"," "),"Action:"," "),"Select application:"," "),"Select the appropriate role:"," "))
     objPDHQ.ActiveWorkbook.Save
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

 ElseIf(objExcel.Cells(rows,1).Value =  "HOME DIRECTORY - CREATION [Add]" OR objExcel.Cells(rows,1).Value =  "HOME DIRECTORY - CREATION [New]") Then
  objSheete.Cells(rows, 1).Value = objExcel.Cells(rows,1).Value
  objSheete.Cells(rows, 2).Value = objSheet.Cells(rows, 2).Value
  objesetup.ActiveWorkbook.Save
  IE.Visible = 1
  Set objHome = CreateObject("Excel.Application")
  objHome.WorkBooks.Open strHome
  Set objSheet2 = objHome.ActiveWorkbook.Worksheets(1)
  IE.Navigate objSheet.Cells(rows, 2).Value
  While IE.ReadyState <> 4
    WScript.Sleep 1000
  Wend
  set urls = ie.document.all.tags("a")
  col = 1
  WScript.Sleep 1000
  Do Until words > 150 OR IE.document.getElementsByTagName("font").Item(words).InnerText =  "-"
   If col = 4 Then
    objSheet2.Cells(row, col).Value =Mid(urls(2).innerHTML,7,9)
    objHome.ActiveWorkbook.Save
    col = col+1
   Else
    If(words < 4 OR words > 10) Then
     objSheet2.Cells(row, col).Value = Trim(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(IE.document.getElementsByTagName("font").Item(words).InnerText,"<"," "),">"," "),"(UID="," "),")", " "),"option:"," "),"Action:"," "),"Select application:"," "),"Select the appropriate role:"," "),"= (", " "))
     objHome.ActiveWorkbook.Save
     col = col+1
    End If
    words= words +1
   End If
  Loop
  row=row+1
  objHome.ActiveWorkbook.Save
  objHome.ActiveWorkbook.Close
  objHome.Application.Quit
  objHome.Quit

 ElseIf(Left(objExcel.Cells(rows,1).Value, 17) =  "DOCIT - QSD [Add]") OR (Left(objExcel.Cells(rows,1).Value, 17) =  "DOCIT - QSD [New]")Then
  objSheete.Cells(rows, 1).Value = Left(objExcel.Cells(rows,1).Value, 17)
  objSheete.Cells(rows, 2).Value = objSheet.Cells(rows, 2).Value
  objesetup.ActiveWorkbook.Save
  IE.Visible = 1
  Set objdocit = CreateObject("Excel.Application")
  objdocit.WorkBooks.Open strdocit
  Set objSheet2 = objdocit.ActiveWorkbook.Worksheets(1)
  IE.Navigate objSheet.Cells(rows, 2).Value
  While IE.ReadyState <> 4
    WScript.Sleep 1000
  Wend
  set urls = ie.document.all.tags("a")
  col = 1
  WScript.Sleep 1000
  Do Until words > 150 OR IE.document.getElementsByTagName("font").Item(words).InnerText =  "-"
   If col = 4 Then
    objSheet2.Cells(rowd, col).Value =Mid(urls(2).innerHTML,7,9)
    objdocit.ActiveWorkbook.Save
    col = col+1
   Else
    If(words < 4 OR words > 10) Then
     objSheet2.Cells(rowd, col).Value = Trim(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(IE.document.getElementsByTagName("font").Item(words).InnerText,"<"," "),">"," "),"(UID="," "),")", " "),"option:"," "),"Action:"," "),"Select application:"," "),"Roles:"," "))
     objdocit.ActiveWorkbook.Save
     col = col+1
    End If
    words= words +1
   End If
  Loop
  rowd=rowd+1
  objdocit.ActiveWorkbook.Save
  objdocit.ActiveWorkbook.Close
  objdocit.Application.Quit
  objdocit.Quit

 ElseIf(objExcel.Cells(rows,1).Value =  "DEVLIMS [Add]") OR (objExcel.Cells(rows,1).Value =  "DEVLIMS [New]")Then
  objSheete.Cells(rows, 1).Value = objExcel.Cells(rows,1).Value
  objSheete.Cells(rows, 2).Value = objSheet.Cells(rows, 2).Value
  objesetup.ActiveWorkbook.Save
  IE.Visible = 1
  Set objdevlim = CreateObject("Excel.Application")
  objdevlim.WorkBooks.Open strdevlim
  Set objSheet2 = objdevlim.ActiveWorkbook.Worksheets(1)
  IE.Navigate objSheet.Cells(rows, 2).Value
  While IE.ReadyState <> 4
    WScript.Sleep 1000
  Wend
  set urls = ie.document.all.tags("a")
  col = 1
  WScript.Sleep 1000
  Do Until words > 150 OR IE.document.getElementsByTagName("font").Item(words).InnerText =  "-"
   If col = 4 Then
    objSheet2.Cells(rowde, col).Value =Mid(urls(2).innerHTML,7,9)
    objdevlim.ActiveWorkbook.Save
    col = col+1
   Else
    If(words < 4 OR words > 10) Then
     objSheet2.Cells(rowde, col).Value = Trim(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(IE.document.getElementsByTagName("font").Item(words).InnerText,"<"," "),">"," "),"(UID="," "),")", " "),"Site:"," "),"Action:"," "),"Environment:"," "),"Roles:"," "))
     objdevlim.ActiveWorkbook.Save
     col = col+1
    End If
    words= words +1
   End If
  Loop
  rowde=rowde+1
  objdevlim.ActiveWorkbook.Save
  objdevlim.ActiveWorkbook.Close
  objdevlim.Application.Quit
  objdevlim.Quit

 ElseIf(Left(objExcel.Cells(rows,1).Value, 25) =  "DIGITAL CERTIFICATE [Add]") OR (Left(objExcel.Cells(rows,1).Value, 25) =  "DIGITAL CERTIFICATE [New]")Then
  objSheete.Cells(rows, 1).Value = Left(objExcel.Cells(rows,1).Value, 25)
  objSheete.Cells(rows, 2).Value = objSheet.Cells(rows, 2).Value
  objesetup.ActiveWorkbook.Save
  IE.Visible = 1
  Set objdigi = CreateObject("Excel.Application")
  objdigi.WorkBooks.Open strdigi
  Set objSheet2 = objdigi.ActiveWorkbook.Worksheets(1)
  IE.Navigate objSheet.Cells(rows, 2).Value
  While IE.ReadyState <> 4
    WScript.Sleep 1000
  Wend
  set urls = ie.document.all.tags("a")
  col = 1
  WScript.Sleep 1000
  Do Until words > 150 OR IE.document.getElementsByTagName("font").Item(words).InnerText =  "-"
   If col = 4 Then
    objSheet2.Cells(rowdi, col).Value =Mid(urls(2).innerHTML,7,9)
    objdigi.ActiveWorkbook.Save
    col = col+1
   Else
    If(words < 4 OR words > 9) Then
     objSheet2.Cells(rowdi, col).Value = Trim(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(IE.document.getElementsByTagName("font").Item(words).InnerText,"<"," "),">"," "),"(UID="," "),")", " "),"Site:"," "),"="," "),"option:"," "),"Digital Certificate Type:"," "))
     objdigi.ActiveWorkbook.Save
     col=col+1
    End If
    If(words=18) Then
     col=col-1
     objSheet2.Cells(rowdi, col).Value = " "
     objdigi.ActiveWorkbook.Save
    End If
    words= words +1
   End If
  Loop
  rowdi=rowdi+1
  objdigi.ActiveWorkbook.Save
  objdigi.ActiveWorkbook.Close
  objdigi.Application.Quit
  objdigi.Quit

 ElseIf(Left(objExcel.Cells(rows,1).Value, 27) =  "TABLEAU SOFTWARE - CONSUMER") Then
  
  IE.Visible = 1
  Set objtab = CreateObject("Excel.Application")
  objtab.WorkBooks.Open strtab
  Set objSheet2 = objtab.ActiveWorkbook.Worksheets(1)
  IE.Navigate objSheet.Cells(rows, 2).Value
  While IE.ReadyState <> 4
    WScript.Sleep 1000
  Wend
  set urls = ie.document.all.tags("a")
  col = 1
  WScript.Sleep 1000
  Do Until words > 150 OR IE.document.getElementsByTagName("font").Item(words).InnerText =  "-"
   If col = 4 Then
    objSheet2.Cells(rowtab, col).Value =Mid(urls(2).innerHTML,7,9)
    objtab.ActiveWorkbook.Save
    col = col+1
   Else
    If(words < 4 OR words > 9) Then
     objSheet2.Cells(rowtab, col).Value = Trim(Replace(Replace(Replace(Replace(Replace(Replace(Replace(IE.document.getElementsByTagName("font").Item(words).InnerText,"<"," "),">"," "),"(UID="," "),")", " "),"Site:"," "),"Action:"," "),"Select a Project: "," "))
     objtab.ActiveWorkbook.Save
     col = col+1
    End If
    words= words +1
   End If
  Loop
  rowtab=rowtab+1
  objtab.ActiveWorkbook.Save
  objtab.ActiveWorkbook.Close
  objtab.Application.Quit
  objtab.Quit

 ElseIf(Left(objExcel.Cells(rows,1).Value, 19) =  "NBR RTREPORTS [Add]") Then
  objSheete.Cells(rows, 1).Value = Left(objExcel.Cells(rows,1).Value, 19)
  objSheete.Cells(rows, 2).Value = objSheet.Cells(rows, 2).Value
  objesetup.ActiveWorkbook.Save
  IE.Visible = 1
  Set objnbr = CreateObject("Excel.Application")
  objnbr.WorkBooks.Open strnbr
  Set objSheet2 = objnbr.ActiveWorkbook.Worksheets(1)
  IE.Navigate objSheet.Cells(rows, 2).Value
  While IE.ReadyState <> 4
    WScript.Sleep 1000
  Wend
  set urls = ie.document.all.tags("a")
  col = 1
  WScript.Sleep 1000
  Do Until words > 150 OR IE.document.getElementsByTagName("font").Item(words).InnerText =  "-"
   If col = 4 Then
    objSheet2.Cells(rownbr, col).Value =Mid(urls(2).innerHTML,7,9)
    objnbr.ActiveWorkbook.Save
    col = col+1
   Else
    If(words < 4 OR words > 9) Then
     objSheet2.Cells(rownbr, col).Value = Trim(Replace(Replace(Replace(Replace(Replace(Replace(Replace(IE.document.getElementsByTagName("font").Item(words).InnerText,"<"," "),">"," "),"(UID="," "),")", " "),"Site:"," "),"option:"," "),"Select Group/Role"," "))
     objnbr.ActiveWorkbook.Save
     col = col+1
    End If
    words= words +1
   End If
  Loop
  rownbr=rownbr+1
  objnbr.ActiveWorkbook.Save
  objnbr.ActiveWorkbook.Close
  objnbr.Application.Quit
  objnbr.Quit

 ElseIf(objExcel.Cells(rows,1).Value =  "AUTO DEPARTURES [Delete]" OR objExcel.Cells(rows,1).Value =  "DEPARTURES [Delete]" OR objExcel.Cells(rows,1).Value =  "DEPARTURES - AUTO [Delete]" OR objExcel.Cells(rows,1).Value =  "DEPARTURE - BP [Delete]" OR objExcel.Cells(rows,1).Value =  "DEPARTURE AUTO - ECLIPSE CLEANUP [Delete]") Then
  objSheete.Cells(rows, 1).Value = objExcel.Cells(rows,1).Value
  objSheete.Cells(rows, 2).Value = objSheet.Cells(rows, 2).Value
  objesetup.ActiveWorkbook.Save
  IE.Visible = 1
  Set objdep = CreateObject("Excel.Application")
  objdep.WorkBooks.Open strdep
  Set objSheet2 = objdep.ActiveWorkbook.Worksheets(1)
  IE.Navigate objSheet.Cells(rows, 2).Value
  While IE.ReadyState <> 4
    WScript.Sleep 1000
  Wend
  set urls = ie.document.all.tags("a")
  col = 1
  WScript.Sleep 1000
  Do Until words > 150 OR IE.document.getElementsByTagName("font").Item(words).InnerText =  "-"
   If col = 4 Then
    objSheet2.Cells(rowdep, col).Value =Mid(urls(2).innerHTML,7,9)
    objdep.ActiveWorkbook.Save
    col = col+1
   Else
    If(words < 4 OR words > 10) Then
     objSheet2.Cells(rowdep, col).Value = Trim(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(IE.document.getElementsByTagName("font").Item(words).InnerText,"<"," "),">"," "),"(UID="," "),")", " "),"option:"," "),"Action:"," "),"Select application:"," "),"Select the appropriate role:"," "))
     objdep.ActiveWorkbook.Save
     col = col+1
    End If
    words= words +1
   End If
  Loop
  rowdep=rowdep+1
  objdep.ActiveWorkbook.Save
  objdep.ActiveWorkbook.Close
  objdep.Application.Quit
  objdep.Quit
  
 ElseIf(Left(objExcel.Cells(rows,1).Value, 10) =  "PRISM/CARA") Then
  objSheete.Cells(rows, 1).Value = Left(objExcel.Cells(rows,1).Value, 10)
  objSheete.Cells(rows, 2).Value = objSheet.Cells(rows, 2).Value
  objesetup.ActiveWorkbook.Save
  IE.Visible = 1
  Set objprism = CreateObject("Excel.Application")
  objprism.WorkBooks.Open strprism
  Set objSheet2 = objprism.ActiveWorkbook.Worksheets(1)
  IE.Navigate objSheet.Cells(rows, 2).Value
  While IE.ReadyState <> 4
    WScript.Sleep 1000
  Wend
  set urls = ie.document.all.tags("a")
  col = 1
  WScript.Sleep 1000
  Do Until words > 150 OR IE.document.getElementsByTagName("font").Item(words).InnerText =  "-"
   If col = 4 Then
    objSheet2.Cells(rowprism, col).Value =Mid(urls(2).innerHTML,7,9)
    objprism.ActiveWorkbook.Save
    col = col+1
   Else
    If(words < 4 OR words > 10) Then
     objSheet2.Cells(rowprism, col).Value = Trim(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(IE.document.getElementsByTagName("font").Item(words).InnerText,"<"," "),">"," "),"(UID="," "),")", " "),"Select role/option:"," "),"Action:"," "),"Select application:"," "),"Select the appropriate role:"," "))
     objprism.ActiveWorkbook.Save
     col = col+1
    End If
    words= words +1
   End If
  Loop
  rowprism=rowprism+1
  objprism.ActiveWorkbook.Save
  objprism.ActiveWorkbook.Close
  objprism.Application.Quit
  objprism.Quit

 ElseIf(Left(objExcel.Cells(rows,1).Value, 9) =  "PPM [Add]") OR (Left(objExcel.Cells(rows,1).Value, 9) =  "PPM [New]")Then
  objSheete.Cells(rows, 1).Value = Left(objExcel.Cells(rows,1).Value, 9)
  objSheete.Cells(rows, 2).Value = objSheet.Cells(rows, 2).Value
  objesetup.ActiveWorkbook.Save
  IE.Visible = 1
  Set objppm = CreateObject("Excel.Application")
  objppm.WorkBooks.Open strppm
  Set objSheet2 = objppm.ActiveWorkbook.Worksheets(1)
  IE.Navigate objSheet.Cells(rows, 2).Value
  While IE.ReadyState <> 4
    WScript.Sleep 1000
  Wend
  set urls = ie.document.all.tags("a")
  col = 1
  WScript.Sleep 1000
  Do Until words > 150 OR IE.document.getElementsByTagName("font").Item(words).InnerText =  "-"
   If col = 4 Then
    objSheet2.Cells(rowppm, col).Value =Mid(urls(2).innerHTML,7,9)
    objppm.ActiveWorkbook.Save
    col = col+1
   Else
    If(words < 4 OR words > 10) Then
     objSheet2.Cells(rowppm, col).Value = Trim(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(IE.document.getElementsByTagName("font").Item(words).InnerText,"<"," "),">"," "),"(UID="," "),")", " "),"Select role/option:"," "),"Action:"," "),"MS Project Pro Needed?:"," "),"Role:"," "))
     objppm.ActiveWorkbook.Save
     col = col+1
    End If
    words= words +1
   End If
  Loop
  rowppm=rowppm+1
  objppm.ActiveWorkbook.Save
  objppm.ActiveWorkbook.Close
  objppm.Application.Quit
  objppm.Quit
 ElseIf(Left(objExcel.Cells(rows,1).Value, 40) =  "TAO / EDC (PRODUCTION - OCP1) [Add Role]" OR Left(objExcel.Cells(rows,1).Value, 40) = "TAO / EDC (TEST/DEV/TRAINING) [Add Role]" ) Then
  objSheete.Cells(rows, 1).Value = Left(objExcel.Cells(rows,1).Value, 40)
  objSheete.Cells(rows, 2).Value = objSheet.Cells(rows, 2).Value
  objesetup.ActiveWorkbook.Save
  IE.Visible = 1
  Set objtao = CreateObject("Excel.Application")
  objtao.WorkBooks.Open strtao
  Set objSheet2 = objtao.ActiveWorkbook.Worksheets(1)
  IE.Navigate objSheet.Cells(rows, 2).Value
  While IE.ReadyState <> 4
    WScript.Sleep 1000
  Wend
  set urls = ie.document.all.tags("a")
  col = 1
  WScript.Sleep 1000
  Do Until words > 150 OR IE.document.getElementsByTagName("font").Item(words).InnerText =  "-"
   If col = 4 Then
    objSheet2.Cells(rowtao, col).Value =Mid(urls(2).innerHTML,7,9)
    objtao.ActiveWorkbook.Save
    col = col+1
   Else
    If(words < 4 OR words > 9) Then
     objSheet2.Cells(rowtao, col).Value = Trim(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(IE.document.getElementsByTagName("font").Item(words).InnerText,"<"," "),">"," "),"(UID="," "),")", " "),"Select role/option:"," "),"Specify Role (EDC Test Only :"," "),"Specify Role (EDC :"," "),"Role:"," "))
     objtao.ActiveWorkbook.Save
     col = col+1
    End If
    words= words +1
   End If
  Loop
  rowtao=rowtao+1
  objtao.ActiveWorkbook.Save
  objtao.ActiveWorkbook.Close
  objtao.Application.Quit
  objtao.Quit
  
 ElseIf(Left(objExcel.Cells(rows,1).Value, 27) =  "SPOTFIRE APPLICATIONS [Add]") OR (Left(objExcel.Cells(rows,1).Value, 27) =  "SPOTFIRE APPLICATIONS [New]")Then
  objSheete.Cells(rows, 1).Value = Left(objExcel.Cells(rows,1).Value, 27)
  objSheete.Cells(rows, 2).Value = objSheet.Cells(rows, 2).Value
  objesetup.ActiveWorkbook.Save
  IE.Visible = 1
  Set objspot = CreateObject("Excel.Application")
  objspot.WorkBooks.Open strspot
  Set objSheet2 = objspot.ActiveWorkbook.Worksheets(1)
  IE.Navigate objSheet.Cells(rows, 2).Value
  While IE.ReadyState <> 4
    WScript.Sleep 1000
  Wend
  set urls = ie.document.all.tags("a")
  col = 1
  WScript.Sleep 1000
  Do Until words > 150 OR IE.document.getElementsByTagName("font").Item(words).InnerText =  "-"
   If col = 4 Then
    objSheet2.Cells(rowspot, col).Value =Mid(urls(2).innerHTML,7,9)
    objspot.ActiveWorkbook.Save
    col = col+1
   Else
    If(words < 4 OR words > 10) Then
     objSheet2.Cells(rowspot, col).Value = Trim(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(IE.document.getElementsByTagName("font").Item(words).InnerText,"<"," "),">"," "),"(UID="," "),")", " "),"Select role/option:"," "),"Action:"," "),"Select application:"," "),"Role:"," "))
     objspot.ActiveWorkbook.Save
     col = col+1
    End If
    words= words +1
   End If
  Loop
  rowspot=rowspot+1
  objspot.ActiveWorkbook.Save
  objspot.ActiveWorkbook.Close
  objspot.Application.Quit
  objspot.Quit
  
 ElseIf(Left(objExcel.Cells(rows,1).Value, 25) =  "PLEXUS FOR PARTNERS [Add]") Then
  objSheete.Cells(rows, 1).Value = Left(objExcel.Cells(rows,1).Value, 25)
  objSheete.Cells(rows, 2).Value = objSheet.Cells(rows, 2).Value
  objesetup.ActiveWorkbook.Save
  IE.Visible = 1
  Set objplex = CreateObject("Excel.Application")
  objplex.WorkBooks.Open strplex
  Set objSheet2 = objplex.ActiveWorkbook.Worksheets(1)
  IE.Navigate objSheet.Cells(rows, 2).Value
  While IE.ReadyState <> 4
    WScript.Sleep 1000
  Wend
  set urls = ie.document.all.tags("a")
  col = 1
  WScript.Sleep 1000
  Do Until words > 150 OR IE.document.getElementsByTagName("font").Item(words).InnerText =  "-"
   If col = 4 Then
    objSheet2.Cells(rowplex, col).Value =Mid(urls(2).innerHTML,7,9)
    objplex.ActiveWorkbook.Save
    col = col+1
   Else
    If(words < 4 OR words > 10) Then
     objSheet2.Cells(rowplex, col).Value = Trim(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(IE.document.getElementsByTagName("font").Item(words).InnerText,"<"," "),">"," "),"(UID="," "),")", " "),"Select role/option:"," "),"Action:"," "),"Alliance Groups: (hold CTRL to multi-select"," "),"Role:"," "))
     objplex.ActiveWorkbook.Save
     col = col+1
    End If
    words= words +1
   End If
  Loop
  rowplex=rowplex+1
  objplex.ActiveWorkbook.Save
  objplex.ActiveWorkbook.Close
  objplex.Application.Quit
  objplex.Quit

 ElseIf(Left(objExcel.Cells(rows,1).Value, 31) =  "JREVIEW (PRODUCTION) [Add Role]" Or Left(objExcel.Cells(rows,1).Value, 38) = "JREVIEW (TEST/DEV/TRAINING) [Add Role]") Then
  objSheete.Cells(rows, 1).Value = Left(objExcel.Cells(rows,1).Value, 31)
  objSheete.Cells(rows, 2).Value = objSheet.Cells(rows, 2).Value
  objesetup.ActiveWorkbook.Save
  IE.Visible = 1
  Set objjr = CreateObject("Excel.Application")
  objjr.WorkBooks.Open strjr
  Set objSheet2 = objjr.ActiveWorkbook.Worksheets(1)
  IE.Navigate objSheet.Cells(rows, 2).Value
  While IE.ReadyState <> 4
    WScript.Sleep 1000
  Wend
  set urls = ie.document.all.tags("a")
  col = 1
  WScript.Sleep 1000
  Do Until words > 150 OR IE.document.getElementsByTagName("font").Item(words).InnerText =  "-"
   If col = 4 Then
    objSheet2.Cells(rowjr, col).Value =Mid(urls(2).innerHTML,7,9)
    objjr.ActiveWorkbook.Save
    col = col+1
   Else
    If(words < 4 OR words > 9) Then
     objSheet2.Cells(rowjr, col).Value = Trim(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(IE.document.getElementsByTagName("font").Item(words).InnerText,"<"," "),">"," "),"(UID="," "),")", " "),"option:"," "),"Action:"," "),"Specify"," "),"Role:"," "))
     objjr.ActiveWorkbook.Save
     col = col+1
    End If
    words= words +1
   End If
  Loop
  rowjr=rowjr+1
  objjr.ActiveWorkbook.Save
  objjr.ActiveWorkbook.Close
  objjr.Application.Quit
  objjr.Quit
  
 ElseIf(Left(objExcel.Cells(rows,1).Value, 21) =  "DIMS - CLINICAL [Add]" Or Left(objExcel.Cells(rows,1).Value, 19) = "DIMS - SAFETY [Add]") Then
  objSheete.Cells(rows, 1).Value = Left(objExcel.Cells(rows,1).Value, 19)
  objSheete.Cells(rows, 2).Value = objSheet.Cells(rows, 2).Value
  objesetup.ActiveWorkbook.Save
  IE.Visible = 1
  Set objdim = CreateObject("Excel.Application")
  objdim.WorkBooks.Open strdim
  Set objSheet2 = objdim.ActiveWorkbook.Worksheets(1)
  IE.Navigate objSheet.Cells(rows, 2).Value
  While IE.ReadyState <> 4
    WScript.Sleep 1000
  Wend
  set urls = ie.document.all.tags("a")
  col = 1
  WScript.Sleep 1000
  Do Until words > 150 OR IE.document.getElementsByTagName("font").Item(words).InnerText =  "-"
   If col = 4 Then
    objSheet2.Cells(rowdim, col).Value =Mid(urls(2).innerHTML,7,9)
    objdim.ActiveWorkbook.Save
    col = col+1
   Else
    If(words < 4 OR words > 9) Then
     objSheet2.Cells(rowdim, col).Value = Trim(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(IE.document.getElementsByTagName("font").Item(words).InnerText,"<"," "),">"," "),"(UID="," "),")", " "),"country:"," "),"Action:"," "),"Specify"," "),"Role:"," "))
     objdim.ActiveWorkbook.Save
     col = col+1
    End If
    words= words +1
   End If
  Loop
  rowdim=rowdim+1
  objdim.ActiveWorkbook.Save
  objdim.ActiveWorkbook.Close
  objdim.Application.Quit
  objdim.Quit
  
 ElseIf(Left(objExcel.Cells(rows,1).Value, 27) =  "CISCO IP COMMUNICATOR [Add]") Then
  objSheete.Cells(rows, 1).Value = Left(objExcel.Cells(rows,1).Value, 27)
  objSheete.Cells(rows, 2).Value = objSheet.Cells(rows, 2).Value
  objesetup.ActiveWorkbook.Save
  IE.Visible = 1
  Set objcisco = CreateObject("Excel.Application")
  objcisco.WorkBooks.Open strcisco
  Set objSheet2 = objcisco.ActiveWorkbook.Worksheets(1)
  IE.Navigate objSheet.Cells(rows, 2).Value
  While IE.ReadyState <> 4
    WScript.Sleep 1000
  Wend
  set urls = ie.document.all.tags("a")
  col = 1
  WScript.Sleep 1000
  Do Until words > 150 OR IE.document.getElementsByTagName("font").Item(words).InnerText =  "-"
   If col = 4 Then
    objSheet2.Cells(rowcisco, col).Value =Mid(urls(2).innerHTML,7,9)
    objcisco.ActiveWorkbook.Save
    col = col+1
   Else
    If(words < 4 OR words > 9) Then
     objSheet2.Cells(rowcisco, col).Value = Trim(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(IE.document.getElementsByTagName("font").Item(words).InnerText,"<"," "),">"," "),"(UID="," "),")", " "),"country:"," "),"Action:"," "),"Specify"," "),"Role:"," "))
     objcisco.ActiveWorkbook.Save
     col = col+1
    End If
    words= words +1
   End If
  Loop
  rowcisco=rowcisco+1
  objcisco.ActiveWorkbook.Save
  objcisco.ActiveWorkbook.Close
  objcisco.Application.Quit
  objcisco.Quit
 
 ElseIf(Left(objExcel.Cells(rows,1).Value, 36) =  "GLOBAL TESTING STANDARDS (GTS) [Add]") Then
  objSheete.Cells(rows, 1).Value = Left(objExcel.Cells(rows,1).Value, 36)
  objSheete.Cells(rows, 2).Value = objSheet.Cells(rows, 2).Value
  objesetup.ActiveWorkbook.Save
  IE.Visible = 1
  Set objgts = CreateObject("Excel.Application")
  objgts.WorkBooks.Open strgts
  Set objSheet2 = objgts.ActiveWorkbook.Worksheets(1)
  IE.Navigate objSheet.Cells(rows, 2).Value
  While IE.ReadyState <> 4
    WScript.Sleep 1000
  Wend
  set urls = ie.document.all.tags("a")
  col = 1
  WScript.Sleep 1000
  Do Until words > 150 OR IE.document.getElementsByTagName("font").Item(words).InnerText =  "-"
   If col = 4 Then
    objSheet2.Cells(rowgts, col).Value =Mid(urls(2).innerHTML,7,9)
    objgts.ActiveWorkbook.Save
    col = col+1
   Else
    If(words < 4 OR words > 9) Then
     objSheet2.Cells(rowgts, col).Value = Trim(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(IE.document.getElementsByTagName("font").Item(words).InnerText,"<"," "),">"," "),"(UID="," "),")", " "),"Action:"," "),"Division:"," "),"Select"," "),"Role:"," "))
     objgts.ActiveWorkbook.Save
     col = col+1
    End If
    words= words +1
   End If
  Loop
  rowgts=rowgts+1
  objgts.ActiveWorkbook.Save
  objgts.ActiveWorkbook.Close
  objgts.Application.Quit
  objgts.Quit 
  
 ElseIf(Left(objExcel.Cells(rows,1).Value, 50) =  "TCANA - TRANSPARENCY CONSENT MANAGEMENT TOOL [Add]") Then
  objSheete.Cells(rows, 1).Value = Left(objExcel.Cells(rows,1).Value, 50)
  objSheete.Cells(rows, 2).Value = objSheet.Cells(rows, 2).Value
  objesetup.ActiveWorkbook.Save
  IE.Visible = 1
  Set objtcana = CreateObject("Excel.Application")
  objtcana.WorkBooks.Open strtcana
  Set objSheet2 = objtcana.ActiveWorkbook.Worksheets(1)
  IE.Navigate objSheet.Cells(rows, 2).Value
  While IE.ReadyState <> 4
    WScript.Sleep 1000
  Wend
  set urls = ie.document.all.tags("a")
  col = 1
  WScript.Sleep 1000
  Do Until words > 150 OR IE.document.getElementsByTagName("font").Item(words).InnerText =  "-"
   If col = 4 Then
    objSheet2.Cells(rowtcana, col).Value =Mid(urls(2).innerHTML,7,9)
    objtcana.ActiveWorkbook.Save
    col = col+1
   Else
    If(words < 4 OR words > 9) Then
     objSheet2.Cells(rowtcana, col).Value = Trim(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(IE.document.getElementsByTagName("font").Item(words).InnerText,"<"," "),">"," "),"(UID="," "),")", " "),"Action:"," "),"Country:"," "),"Select"," "),"Role:"," "))
     objtcana.ActiveWorkbook.Save
     col = col+1
    End If
    words= words +1
   End If
  Loop
  rowtcana=rowtcana+1
  objtcana.ActiveWorkbook.Save
  objtcana.ActiveWorkbook.Close
  objtcana.Application.Quit
  objtcana.Quit
  
 ElseIf(Left(objExcel.Cells(rows,1).Value, 27) =  "AUTODESK APPLICATIONS [Add]") Then
  objSheete.Cells(rows, 1).Value = Left(objExcel.Cells(rows,1).Value, 50)
  objSheete.Cells(rows, 2).Value = objSheet.Cells(rows, 2).Value
  objesetup.ActiveWorkbook.Save
  IE.Visible = 1
  Set objautodsk = CreateObject("Excel.Application")
  objautodsk.WorkBooks.Open strautodsk
  Set objSheet2 = objautodsk.ActiveWorkbook.Worksheets(1)
  IE.Navigate objSheet.Cells(rows, 2).Value
  While IE.ReadyState <> 4
    WScript.Sleep 1000
  Wend
  set urls = ie.document.all.tags("a")
  col = 1
  WScript.Sleep 1000
  Do Until words > 150 OR IE.document.getElementsByTagName("font").Item(words).InnerText =  "-"
   If col = 4 Then
    objSheet2.Cells(rowautodsk, col).Value =Mid(urls(2).innerHTML,7,9)
    objautodsk.ActiveWorkbook.Save
    col = col+1
   Else
    If(words < 4 OR words > 9) Then
     objSheet2.Cells(rowautodsk, col).Value = Trim(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(IE.document.getElementsByTagName("font").Item(words).InnerText,"<"," "),">"," "),"(UID="," "),")", " "),"Action:"," "),"Country:"," "),"Application(s :"," "),"Role:"," "))
     objautodsk.ActiveWorkbook.Save
     col = col+1
    End If
    words= words +1
   End If
  Loop
  rowautodsk=rowautodsk+1
  objautodsk.ActiveWorkbook.Save
  objautodsk.ActiveWorkbook.Close
  objautodsk.Application.Quit
  objautodsk.Quit
 
 ElseIf(Left(objExcel.Cells(rows,1).Value, 39) =  "DELTAV RECIPE REVIEW SYSTEM (RRS) [Add]") Then
  objSheete.Cells(rows, 1).Value = Left(objExcel.Cells(rows,1).Value, 50)
  objSheete.Cells(rows, 2).Value = objSheet.Cells(rows, 2).Value
  objesetup.ActiveWorkbook.Save
  IE.Visible = 1
  Set objdeltav = CreateObject("Excel.Application")
  objdeltav.WorkBooks.Open strdeltav
  Set objSheet2 = objdeltav.ActiveWorkbook.Worksheets(1)
  IE.Navigate objSheet.Cells(rows, 2).Value
  While IE.ReadyState <> 4
    WScript.Sleep 1000
  Wend
  set urls = ie.document.all.tags("a")
  col = 1
  WScript.Sleep 1000
  Do Until words > 150 OR IE.document.getElementsByTagName("font").Item(words).InnerText =  "-"
   If col = 4 Then
    objSheet2.Cells(rowdeltav, col).Value =Mid(urls(2).innerHTML,7,9)
    objdeltav.ActiveWorkbook.Save
    col = col+1
   Else
    If(words < 4 OR words > 9) Then
     objSheet2.Cells(rowdeltav, col).Value = Trim(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(IE.document.getElementsByTagName("font").Item(words).InnerText,"<"," "),">"," "),"(UID="," "),")", " "),"Action:"," "),"Country:"," "),"Type:"," "),"Role:"," "))
     objdeltav.ActiveWorkbook.Save
     col = col+1
    End If
    words= words +1
   End If
  Loop
  rowdeltav=rowdeltav+1
  objdeltav.ActiveWorkbook.Save
  objdeltav.ActiveWorkbook.Close
  objdeltav.Application.Quit
  objdeltav.Quit
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
