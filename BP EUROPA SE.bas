Attribute VB_Name = "Module1"
Public Sub automatenettingBPEUROPA()
    Dim sapGuiApp, sapGuiAuto, xclapp, xclwbk, connection, session, WScrip  As Object
    Dim lastRow, lastRow1, newRow, i, newRow1 As Long
    Dim rowIndex, ExcelRow As Integer
    Dim excelApp As Excel.Application
    Dim excelWorkbook As Excel.Workbook
    Dim excelWorksheet As Excel.Worksheet
    Dim excelFilePath As String
    Dim sourceWorkbook, source1Workbook, destinationWorkbook As Workbook
    Dim destinationWorkbookPath, Fromdate, Todate, Savelocation As String
    Dim destinationWorksheet, wsOriginal As Worksheet
    Dim cell As Range
    Dim startdate, enddate As Date
    Dim sapGuiVersion, sapGuiScriptingVersion, CompanyCode, formatstartdate, formatenddate, Rate, CustomerAccount, Reference As String
    
    
    
     'Created by  Rebecca Guisgo 17-04-2024
      
    
   
   
    
    'connecting to SAP
    
    If Not IsObject(sapGuiApp) Then
       Set sapGuiAuto = GetObject("SAPGUI")
       Set sapGuiApp = sapGuiAuto.GetScriptingEngine
    End If
    If Not IsObject(connection) Then
       Set connection = sapGuiApp.Children(0)
    End If
    If Not IsObject(session) Then
       Set session = connection.Children(0)
    End If
    If IsObject(WScript) Then
       WScript.ConnectObject session, "on"
       WScript.ConnectObject sapGuiApp, "on"
    End If
    
    Savelocation = "I:\Controllers\BSC Budapest\PMI\Protected folders\Exchange Team\NETTING 2024\netting\"
    
    
    startdate = DateSerial(Year(Date), Month(Date), 1)
    enddate = DateSerial(Year(Date), Month(Date) + 1, 0)
    formatstartdate = Format(startdate, "DDMMYYYY")
    formatenddate = Format(enddate, "DDMMYYYY")
    
   
    'select vendor number and company code
    'session.findById("wnd[0]").maximize
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nfbl1"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/ctxtKD_LIFNR-LOW").Text = 554619
    session.findById("wnd[0]/usr/ctxtKD_BUKRS-LOW").Text = 1742
    session.findById("wnd[0]/usr/chkX_APAR").Selected = True
    session.findById("wnd[0]/usr/ctxtPA_VARI").Text = "/netting"
    'session.findById("wnd[0]/usr/radX_AISEL").Select
    'session.findById("wnd[0]/usr/ctxtSO_BUDAT-LOW").Text = "01122023"
    'session.findById("wnd[0]/usr/ctxtSO_BUDAT-HIGH").Text = "20022024"
    session.findById("wnd[0]/usr/ctxtPA_VARI").SetFocus
    session.findById("wnd[0]/usr/ctxtPA_VARI").caretPosition = 8
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    session.findById("wnd[0]/usr/lbl[88,8]").SetFocus
    session.findById("wnd[0]/usr/lbl[88,8]").caretPosition = 0
    session.findById("wnd[0]").sendVKey 2
    session.findById("wnd[0]/tbar[1]/btn[38]").press
    session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").Text = "M"
    session.findById("wnd[1]").sendVKey 0
    session.findById("wnd[0]/usr/lbl[66,8]").SetFocus
    session.findById("wnd[0]/usr/lbl[66,8]").caretPosition = 2
    session.findById("wnd[0]").sendVKey 2
    session.findById("wnd[0]/tbar[1]/btn[38]").press
    session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").Text = formatstartdate
    session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-HIGH").Text = formatenddate
    session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-HIGH").SetFocus
    session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-HIGH").caretPosition = 10
    session.findById("wnd[1]").sendVKey 0
    session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[2]").Select
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
    'session.findById("wnd[1]/tbar[0]/btn[0]").press
    session.findById("wnd[1]").sendVKey 0
    session.findById("wnd[1]/usr/ctxtDY_PATH").Text = Savelocation
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = 554619 & " " & formatstartdate & " " & "export.xls"
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    session.findById("wnd[0]/tbar[0]/btn[3]").press



    Set source1Workbook = Workbooks.Open(Savelocation & 554619 & " " & formatstartdate & " " & "export.xls")
    Range("F:F").Select
    Selection.Delete Shift:=xlLeft
  
     'Prompt to select workbook
    
    
    sourceWorkbookPath = Application.GetOpenFilename(FileFilter:="Excel Files (*.xls; *.xlsx),*.xls; *xlsx", Title:="Select Workbook")
    
    'check workbook is selected
    
    If sourceWorkbook <> "False" Then
        'open selected workbook
        Set sourceWorkbook = Workbooks.Open(sourceWorkbookPath)
    End If
    
    Worksheet = InputBox("please type your sheetname:")
     If Worksheet = "" Then
        MsgBox "sheetname cannot be empty. Macro will exit.", vbExclamation
        Exit Sub
    End If
    Set destinationWorksheet = sourceWorkbook.Sheets(Worksheet)
    
    Set wsOriginal = source1Workbook.Sheets(1)
    
    wsOriginal.Activate
    destinationWorksheet.Activate
   
        
    lastRow = wsOriginal.Cells(wsOriginal.Rows.Count, "G").End(xlUp).Row
        
        'loop through each row in column G
        
   
        
     'check the first  characters
     
    newRow = 11 'start pasting from row 11
    newRow1 = 11
    For i = 1 To lastRow
        If Left(wsOriginal.Cells(i, "G").Value, 1) = "3" Then
            wsOriginal.Rows(i).Columns("D:G").Copy
            destinationWorksheet.Cells(newRow, "B").PasteSpecial Paste:=xlPasteValues
            newRow = newRow + 1
        ElseIf Left(wsOriginal.Cells(i, "G").Value, 1) = "5" Then
             wsOriginal.Rows(i).Columns("D:G").Copy
             destinationWorksheet.Cells(newRow1, "G").PasteSpecial Paste:=xlPasteValues
             newRow1 = newRow1 + 1
        End If
    Next i
    
    Application.CutCopyMode = False
    lastRow1 = destinationWorksheet.Cells(destinationWorksheet.Rows.Count, "I").End(xlUp).Row
     For Each cell In destinationWorksheet.Range("I11:I" & lastRow1 - 1)
         If cell.Value < 0 Then
            cell.Value = Abs(cell.Value)
        ElseIf cell.Value > 0 Then
            cell.Value = -cell.Value
        End If
    Next cell
    
   
   
    
    
End Sub

