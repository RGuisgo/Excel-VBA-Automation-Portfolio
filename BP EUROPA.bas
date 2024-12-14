Attribute VB_Name = "Module1"
Public Sub automatenettingPAYABLEBPEUROPA()
    Dim sapGuiApp, sapGuiAuto, xclapp, xclwbk, connection, session, WScrip  As Object
    Dim lastRow, lastRow1, newRow, i, newRow1 As Long
    Dim rowindex, ExcelRow, rowcount, saprowindex As Integer
    Dim sapGuiVersion, sapGuiScriptingVersion, CompanyCode, VendorAccount, Rate, CustomerAccount, Reference As String
    Dim excelApp As Excel.Application
    Dim excelWorkbook As Excel.Workbook
    Dim excelWorksheet As Excel.Worksheet
    Dim excelFilePath, value As String
    Dim sourceWorkbook, source1Workbook, destinationWorkbook As Workbook
    Dim destinationWorkbookPath, Fromdate, Todate, Docdate, DueDate, Savelocation As String
    Dim destinationWorksheet, wsn, wsOriginal As Worksheet
    Dim cell, items As Range
    Dim formatstartdate, formatenddate As String
    Dim startdate, enddate As Date

    
    
     'Created by  Rebecca Guisgo 08-03-2024
      
    
    Rate = InputBox("Please type in the corresponding Currency:")
    If Rate = "" Then
        MsgBox "Currency cannot be empty. Macro will exit.", vbExclamation
        Exit Sub
    End If
    Reference = InputBox("Please type in the corresponding Reference:")
    If Reference = "" Then
        MsgBox "Reference cannot be empty. Macro will exit.", vbExclamation
        Exit Sub
    End If
    Docdate = InputBox("Please type in the corresponding Docdate:")
    If Docdate = "" Then
        MsgBox "Docdate cannot be empty. Macro will exit.", vbExclamation
        Exit Sub
    End If
    
    DueDate = InputBox("Please type in the corresponding duedate:")
    If DueDate = "" Then
        MsgBox "duedate cannot be empty. Macro will exit.", vbExclamation
        Exit Sub
    End If
    
    
    
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
    session.findById("wnd[0]/usr/ctxtKD_BUKRS-LOW").Text = "1742"
    session.findById("wnd[0]/usr/chkX_APAR").Selected = True
    session.findById("wnd[0]/usr/ctxtPA_VARI").Text = "/netting"
    'session.findById("wnd[0]/usr/radX_AISEL").Select
    'session.findById("wnd[0]/usr/ctxtSO_BUDAT-LOW").Text = "01122023"
    'session.findById("wnd[0]/usr/ctxtSO_BUDAT-HIGH").Text = "20032024"
    session.findById("wnd[0]/usr/ctxtPA_VARI").SetFocus
    session.findById("wnd[0]/usr/ctxtPA_VARI").caretPosition = 8
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    session.findById("wnd[0]/usr/lbl[88,8]").SetFocus ' all other
    session.findById("wnd[0]/usr/lbl[88,8]").caretPosition = 0 ' all other
    'session.findById("wnd[0]/usr/lbl[88,4]").SetFocus ' for scaped
    'session.findById("wnd[0]/usr/lbl[88,4]").caretPosition = 0 'for scaped
    session.findById("wnd[0]").sendVKey 2
    session.findById("wnd[0]/tbar[1]/btn[38]").press
    session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").Text = "M"
    session.findById("wnd[1]").sendVKey 0
    session.findById("wnd[0]/usr/lbl[66,8]").SetFocus ' all other
    session.findById("wnd[0]/usr/lbl[66,8]").caretPosition = 2 'all other
    'session.findById("wnd[0]/usr/lbl[66,4]").SetFocus ' for scaped
    'session.findById("wnd[0]/usr/lbl[66,4]").caretPosition = 2 ' for scaped
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
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = 554619 & " 1 " & formatstartdate & " " & "export.xls"
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    session.findById("wnd[0]/tbar[0]/btn[3]").press

  
    Set source1Workbook = Workbooks.Open(Savelocation & 554619 & " 1 " & formatstartdate & " " & "export.xls")
    Range("F:F").Select
    Selection.Delete Shift:=xlLeft
    
    
   
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
    
    
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nf-51"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/ctxtBKPF-BLDAT").Text = Docdate
    session.findById("wnd[0]/usr/ctxtBKPF-BUKRS").Text = "1742"
    session.findById("wnd[0]/usr/ctxtBKPF-WAERS").Text = Rate
    session.findById("wnd[0]/usr/txtBKPF-XBLNR").Text = Reference
    session.findById("wnd[0]/usr/txtBKPF-XBLNR").SetFocus
    session.findById("wnd[0]/usr/txtBKPF-XBLNR").caretPosition = 14
    session.findById("wnd[0]/tbar[1]/btn[6]").press
    session.findById("wnd[0]/usr/chkRF05A-XMULK").Selected = True
    session.findById("wnd[0]/usr/sub:SAPMF05A:0710/radRF05A-XPOS1[2,0]").Select
    'session.findById("wnd[0]/usr/ctxtRF05A-AGKON").Text = VendorAccount
    session.findById("wnd[0]/usr/sub:SAPMF05A:0710/radRF05A-XPOS1[2,0]").SetFocus
    session.findById("wnd[0]/tbar[1]/btn[16]").press
    session.findById("wnd[1]/usr/sub:SAPMF05A:0609/chkRF05A-XNOPS[0,29]").Selected = True
    session.findById("wnd[1]/usr/sub:SAPMF05A:0609/chkRF05A-XNOPS[1,29]").Selected = True
    'session.findById("wnd[1]/usr/sub:SAPMF05A:0609/ctxtRF05A-AGKON[0,0]").Text = VendorAccount
    session.findById("wnd[1]/usr/sub:SAPMF05A:0609/ctxtRF05A-AGKOA[0,17]").Text = "K"
    session.findById("wnd[1]/usr/sub:SAPMF05A:0609/ctxtRF05A-AGBUK[0,23]").Text = "1742"
    'session.findById("wnd[1]/usr/sub:SAPMF05A:0609/ctxtRF05A-AGKON[1,0]").Text = CustomerAccount
    session.findById("wnd[1]/usr/sub:SAPMF05A:0609/ctxtRF05A-AGKOA[1,17]").Text = "D"
    session.findById("wnd[1]/usr/sub:SAPMF05A:0609/ctxtRF05A-AGBUK[1,23]").Text = "1742"
    session.findById("wnd[1]/usr/sub:SAPMF05A:0609/chkRF05A-XNOPS[1,29]").SetFocus
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    
    Set wsn = source1Workbook.ActiveSheet
    
    rowcount = wsn.Cells(wsn.Rows.Count, "G").End(xlUp).Row
    rowindex = 11 ' check this, differes from RU to RU
    saprowindex = 0
    lastRow = rowcount - 1
    Do While rowindex <= rowcount
        value = wsn.Cells(rowindex, "G").value
        session.findById("wnd[0]/usr/sub:SAPMF05A:0731/txtRF05A-SEL01[" & saprowindex & ",0]").Text = value
        rowindex = rowindex + 1
        saprowindex = saprowindex + 1
        If saprowindex > 15 Then
            session.findById("wnd[0]/usr/sub:SAPMF05A:0731/txtRF05A-SEL01[12,0]").SetFocus
            session.findById("wnd[0]/usr/sub:SAPMF05A:0731/txtRF05A-SEL01[12,0]").caretPosition = 10
            session.findById("wnd[0]/tbar[1]/btn[16]").press
            'session.findById("wnd[1]/tbar[0]/btn[0]").press
            session.findById("wnd[0]/tbar[1]/btn[7]").press
            session.findById("wnd[0]/tbar[1]/btn[6]").press
            session.findById("wnd[0]/usr/sub:SAPMF05A:0710/radRF05A-XPOS1[2,0]").Select
            session.findById("wnd[0]/usr/chkRF05A-XMULK").Selected = True
            session.findById("wnd[0]/usr/chkRF05A-XMULK").SetFocus
            session.findById("wnd[0]/tbar[1]/btn[16]").press
            session.findById("wnd[1]/usr/sub:SAPMF05A:0609/chkRF05A-XNOPS[0,29]").Selected = True
            session.findById("wnd[1]/usr/sub:SAPMF05A:0609/chkRF05A-XNOPS[1,29]").Selected = True
            'session.findById("wnd[1]/usr/sub:SAPMF05A:0609/ctxtRF05A-AGKON[0,0]").Text = CustomerAccount
            session.findById("wnd[1]/usr/sub:SAPMF05A:0609/ctxtRF05A-AGKOA[0,17]").Text = "D"
            session.findById("wnd[1]/usr/sub:SAPMF05A:0609/ctxtRF05A-AGBUK[0,23]").Text = "1742"
           ' session.findById("wnd[1]/usr/sub:SAPMF05A:0609/ctxtRF05A-AGKON[1,0]").Text = VendorAccount
            session.findById("wnd[1]/usr/sub:SAPMF05A:0609/ctxtRF05A-AGKOA[1,17]").Text = "K"
            session.findById("wnd[1]/usr/sub:SAPMF05A:0609/ctxtRF05A-AGBUK[1,23]").Text = "1742"
            session.findById("wnd[1]/usr/sub:SAPMF05A:0609/chkRF05A-XNOPS[1,29]").SetFocus
            session.findById("wnd[1]/tbar[0]/btn[0]").press
            saprowindex = 0
        End If
    Loop
    totalsum = 0
    For i = 11 To rowcount
        If Not IsEmpty(wsn.Cells(i, "G").value) Then
            totalsum = totalsum + wsn.Cells(i, "F").value
        End If
    Next i
    session.findById("wnd[0]/tbar[1]/btn[16]").press
    'session.findById("wnd[1]/tbar[0]/btn[0]").press ' for withholding tax thing, doesnt apply to every RU
    session.findById("wnd[0]/tbar[1]/btn[14]").press
    session.findById("wnd[0]/usr/ctxtRF05A-NEWBS").Text = "37"
    session.findById("wnd[0]/usr/ctxtRF05A-NEWKO").Text = 554619
    session.findById("wnd[0]/usr/ctxtRF05A-NEWKO").SetFocus
    session.findById("wnd[0]/usr/ctxtRF05A-NEWKO").caretPosition = 6
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/txtBSEG-WRBTR").Text = Abs(totalsum)
    session.findById("wnd[0]/usr/ctxtBSEG-GSBER").Text = "7400"
    session.findById("wnd[0]/usr/ctxtBSEG-ZTERM").Text = "Z005"
    session.findById("wnd[0]/usr/ctxtBSEG-ZFBDT").Text = DueDate
    session.findById("wnd[0]/usr/ctxtBSEG-ZLSCH").Text = "x"
    session.findById("wnd[0]/usr/ctxtBSEG-SGTXT").Text = Reference
    session.findById("wnd[0]/usr/ctxtBSEG-SGTXT").SetFocus
    session.findById("wnd[0]/usr/ctxtBSEG-SGTXT").caretPosition = 7
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/tbar[1]/btn[7]").press
    session.findById("wnd[0]/usr/ctxtBSEG-BVTYP").SetFocus
    session.findById("wnd[0]/usr/ctxtBSEG-BVTYP").caretPosition = 0
    session.findById("wnd[0]").sendVKey 4
   ' session.findById("wnd[1]").sendVKey 2 - it ends here where we have to select bank account for payables.
   ' session.findById("wnd[0]").sendVKey 0
   ' session.findById("wnd[1]").sendVKey 0
    
   
    
End Sub




