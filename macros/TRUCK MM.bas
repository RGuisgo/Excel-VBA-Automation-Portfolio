Attribute VB_Name = "mod_MM"
Public Function MM()

If Not IsObject(SAPGUIApp) Then
   Set SapGuiAuto = GetObject("SAPGUI")
   Set SAPGUIApp = SapGuiAuto.GetScriptingEngine
End If

If Not IsObject(Connection) Then
   Set Connection = SAPGUIApp.Children(0)
End If

If Not IsObject(session) Then
   Set session = Connection.Children(0)
End If

If IsObject(WScript) Then
   WScript.ConnectObject session, "on"
   WScript.ConnectObject SAPGUIApp, "on"
End If

Dim strScript As String
Dim wrkbk As Workbook
Dim wrksht As Worksheet
Dim row As Integer
Dim lastRow, i, j, k, saprowIndex, valueindex, rowindex As Long
Dim Plant, Material, materialcode(), key As String
Dim materialplantcodes, codes As Variant
Dim valuecloumnc As Variant
Dim Amount As Variant
Dim response As VbMsgBoxResult



Set wrkbk = Excel.Application.Workbooks(ThisWorkbook.Name)
Set wrksht = wrkbk.Sheets("MM")

row = 2

While wrksht.Cells(row, lookfor(wrksht, "No.")) <> ""

'Filter Doc Number
session.FindById("wnd[0]/usr/cntlSINWP_CONTAINER/shellcont/shell/shellcont[1]/shell/shellcont[0]/shell").SetCurrentCell -1, "WIOBJID"
session.FindById("wnd[0]/usr/cntlSINWP_CONTAINER/shellcont/shell/shellcont[1]/shell/shellcont[0]/shell").SelectColumn "WIOBJID"
session.FindById("wnd[0]/usr/cntlSINWP_CONTAINER/shellcont/shell/shellcont[1]/shell/shellcont[0]/shell").ContextMenu
session.FindById("wnd[0]/usr/cntlSINWP_CONTAINER/shellcont/shell/shellcont[1]/shell/shellcont[0]/shell").SelectContextMenuItem "&FILTER"
session.FindById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").Text = wrksht.Cells(row, lookfor(wrksht, "WIContent"))
session.FindById("wnd[1]/tbar[0]/btn[0]").Press

'Open Doc Number
session.FindById("wnd[0]/usr/cntlSINWP_CONTAINER/shellcont/shell/shellcont[1]/shell/shellcont[0]/shell").CurrentCellColumn = "WIOBJID"
session.FindById("wnd[0]/usr/cntlSINWP_CONTAINER/shellcont/shell/shellcont[1]/shell/shellcont[0]/shell").SelectedRows = "0"
session.FindById("wnd[0]/usr/cntlSINWP_CONTAINER/shellcont/shell/shellcont[1]/shell/shellcont[0]/shell").DoubleClickCurrentCell

'Attachment Pop up
On Error Resume Next
session.FindById("wnd[1]/usr/cntlCUSTOM_CONTAINER_100/shellcont/shell").CurrentCellColumn = "BITM_DESCR"
session.FindById("wnd[1]/usr/cntlCUSTOM_CONTAINER_100/shellcont/shell").SelectedRows = "0"
session.FindById("wnd[1]/usr/cntlCUSTOM_CONTAINER_100/shellcont/shell").DoubleClickCurrentCell
session.FindById("wnd[1]").Close
On Error GoTo 0

'Preventing it does not work with GOM open
'Do Until InStr(1, SAPGUIApp.Children(CLng(ConNum)).Description, "EUP") > 0
'    ConNum = ConNum + 1
'    Loop
'    Set SAPCon = SAPGUIApp.Children(CLng(ConNum))
'    Set session = SAPCon.Children(0)

                                                'VIM Indexing

    'Basic Data Tab
IndexData = "wnd[0]/usr/subSUB_MAIN:/OPT/SAPLVIM_IDX_UI:1001/subSUB_TAB_STRIP:/OPT/SAPLVIM_IDX_UI:1002/tabsTAB_MAIN/"
BasicData = IndexData & "tabpTAB1/ssubTAB_MAIN_SUBSCREEN:SAPLZF_VIM_IDX_UI:8101/ctxtGH_IDX_APPLICATION->MS_IDX_HEADER-"

'Do not auto post
session.FindById(IndexData & "tabpTAB1/ssubTAB_MAIN_SUBSCREEN:SAPLZF_VIM_IDX_UI:8101/chkGH_IDX_APPLICATION->MS_IDX_HEADER-CUSTOM_FIELD4").Selected = True

'Baseline Date
session.FindById(BasicData & "ZZBASELINE_DATE").Text = wrksht.Cells(row, lookfor(wrksht, "BaselineDate"))

'Payment Term
session.FindById(BasicData & "PYMNT_TERMS").Text = "Z005"

'Tax Code
session.FindById(BasicData & "TAX_CODE").Text = "VL"

'Basic Text
session.FindById(IndexData & "tabpTAB1/ssubTAB_MAIN_SUBSCREEN:SAPLZF_VIM_IDX_UI:8101/txtGH_IDX_APPLICATION->MS_IDX_HEADER-SGTXT").Text = wrksht.Cells(row, lookfor(wrksht, "BasicText"))

'Untick Auto Calc Tax Flag
session.FindById(IndexData & "tabpTAB1/ssubTAB_MAIN_SUBSCREEN:SAPLZF_VIM_IDX_UI:8101/chkGH_IDX_APPLICATION->MS_IDX_HEADER-AUTO_CALC").Selected = False

    'Vendor Info Tab
session.FindById(IndexData & "tabpTAB4").Select
VendorInfo = IndexData & "tabpTAB4/ssubTAB_MAIN_SUBSCREEN:SAPLZF_VIM_IDX_UI:8103/ctxtGH_IDX_APPLICATION->MS_IDX_HEADER-"

'Permitted Payee Number
session.FindById(VendorInfo & "ATTRIBUTE1").Text = "0000555446"

'Bank Account
session.FindById(VendorInfo & "BVTYP").Text = "0002"

                                                'MM Invoice
'Run Business Rule - Go to Miro
On Error Resume Next
session.FindById("wnd[0]/usr/subSUB_MAIN:/OPT/SAPLVIM_IDX_UI:1001/subSUB_PROC_OPTIONS:/OPT/SAPLVIM_IDX_UI:1003/cntlCC_PROCESS_OPTIONS/shellcont/shell").CurrentCellRow = 3
session.FindById("wnd[0]/usr/subSUB_MAIN:/OPT/SAPLVIM_IDX_UI:1001/subSUB_PROC_OPTIONS:/OPT/SAPLVIM_IDX_UI:1003/cntlCC_PROCESS_OPTIONS/shellcont/shell").PressButtonCurrentCell
session.FindById("wnd[1]/usr/btnBUTTON_2").Press
On Error GoTo 0

' Remove Bank Account
PaymentTab = "wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_PAY/"
session.FindById(PaymentTab & "ssubHEADER_SCREEN:SAPLFDCB:0020/ctxtINVFO-BVTYP").Text = ""
session.FindById(PaymentTab & "ssubHEADER_SCREEN:SAPLFDCB:0020/ctxtINVFO-BVTYP").CaretPosition = 0
session.FindById("wnd[0]").SendVKey 0

'MIRO
strScript = "wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_TOTAL/ssubHEADER_SCREEN:SAPLFDCB:0010/"

''Invoice Date
'session.FindById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_TOTAL").Select
'InvoiceDate = session.FindById(strScript & "ctxtINVFO-BLDAT").Text
'wrksht.Cells(row, lookfor(wrksht, "InvoiceDate")) = InvoiceDate

'PO Tab
session.FindById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/subITEMS:SAPLMR1M:6010/tabsITEMTAB/tabpITEMS_PO/ssubTABS:SAPLMR1M:6020/subREFERENZBELEG:SAPLMR1M:6211/btnRM08M-XMSEL").Press
session.FindById("wnd[1]/usr/subMSEL:SAPLMR1M:6221/tblSAPLMR1MTC_MSEL_BEST/ctxtRM08M-EBELN[0,0]").SetFocus
session.FindById("wnd[1]/usr/subMSEL:SAPLMR1M:6221/tblSAPLMR1MTC_MSEL_BEST/ctxtRM08M-EBELN[0,0]").CaretPosition = 0
session.FindById("wnd[1]").SendVKey 4

'PO Tab - Vendor Code
session.FindById("wnd[0]/usr/ctxtSO_LIFNR-LOW").Text = wrksht.Cells(row, lookfor(wrksht, "Vendor"))
session.FindById("wnd[0]/usr/ctxtSO_BEDAT-LOW").Text = wrksht.Cells(row, lookfor(wrksht, "StartDate"))
session.FindById("wnd[0]/usr/ctxtSO_BEDAT-HIGH").Text = wrksht.Cells(row, lookfor(wrksht, "EndDate"))


'create a dictionary to map plant codes and material names
Set materialplantcodes = CreateObject("Scripting.Dictionary")

   'Set wrkbk = Excel.Application.Workbooks(ThisWorkbook.Name)
Set wrksht = ThisWorkbook.Sheets("MM")
   
   
    
  ' define the material name, plant code, code mappings
  
  
  'diesel
  
materialplantcodes.Add "Diesel_0553", Array(101415, 152384, 101428)
materialplantcodes.Add "Diesel_0600", Array(101415, 152384, 101428)
materialplantcodes.Add "Diesel_0633", Array(101415, 152384, 101428)
materialplantcodes.Add "Diesel_0639", Array(101415, 152384, 101428)
materialplantcodes.Add "Diesel_0n22", Array(101415, 152384, 101428)
materialplantcodes.Add "Diesel_0598", Array(101415, 152384, 101428)
  
  
  ' Supervol
  
materialplantcodes.Add "SuperVol_0553", Array(101380, 101381)
materialplantcodes.Add "SuperVol_0600", Array(101380, 101381)
materialplantcodes.Add "SuperVol_0633", Array(101380, 101381)
materialplantcodes.Add "SuperVol_0639", Array(101380, 101381)
materialplantcodes.Add "SuperVol_0n22", Array(101380, 101381)
materialplantcodes.Add "SuperVol_0598", Array(101380, 101381)
  
  
  'super plus
  
materialplantcodes.Add "SuperPlus_0553", Array(101387, 101392)
materialplantcodes.Add "SuperPlus_0600", Array(101387, 101392)
materialplantcodes.Add "SuperPlus_0633", Array(101387, 101392)
materialplantcodes.Add "SuperPlus_0639", Array(101387, 101392)
materialplantcodes.Add "SuperPlus_0n22", Array(101387, 101392)
materialplantcodes.Add "SuperPlus_0598", Array(101387, 101392)
  
  'superE10
  
materialplantcodes.Add "SuperE10_0553", Array(151602, 152259)
materialplantcodes.Add "SuperE10_0600", Array(151602, 152259)
materialplantcodes.Add "SuperE10_0633", Array(151602, 152259)
materialplantcodes.Add "SuperE10_0639", Array(151602, 152259)
materialplantcodes.Add "SuperE10_0n22", Array(151602, 152259)
materialplantcodes.Add "SuperE10_0598", Array(151602, 152259)

  
  ' Heizol
  
  
materialplantcodes.Add "Heizol_0553", "150769"
materialplantcodes.Add "Heizol_0600", "150769"
materialplantcodes.Add "Heizol_0633", "150769"
materialplantcodes.Add "Heizol_0639", "150769"
materialplantcodes.Add "Heizol_0n22", "150769"
materialplantcodes.Add "Heizol_0598", "150769"
  
  
  
     
      
  
  'find last row in plant column
    
    
lastRow = wrksht.Cells(wrksht.Rows.Count, "K").End(xlUp).row
  
  
i = 2
Material = wrksht.Cells(i, "L").Value
Plant = wrksht.Cells(i, "M").Value
key = Material & "_" & Plant
      
      
      'get the codes from the dictionary based on material name and plant code
      
codes = GetCodes(materialplantcodes, key)

      ' set plant code in sap
      
      
      'If ws.Cells(i, "L").Value = ws.Cells(i + 1, "L").Value And ws.Cells(i, "K").Value = ws.Cells(i + 1, "K").Value Then
session.FindById("wnd[0]/usr/ctxtSO_WERKS-LOW").Text = Plant
      'Else
      'End If
      'set material name and code
      
k = 0
If IsArray(codes) Then
    'For k = 0 To 2
        session.FindById("wnd[0]/usr/btn%_SO_MATNR_%_APP_%-VALU_PUSH").Press
        For Each code In codes
            session.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1," & k & "]").Text = code
    'session.findById("wnd[0]/usr/ctxtSO_MATNR-LOW").Text = code
    'session.findById("wnd[0]/usr/ctxtSO_MATNR-HIGH").Text = code
            k = k + 1
        Next code
        
    'Next k
    session.FindById("wnd[1]/tbar[0]/btn[8]").Press
    session.FindById("wnd[0]/tbar[1]/btn[8]").Press
    session.FindById("wnd[0]/tbar[1]/btn[5]").Press
    session.FindById("wnd[0]/tbar[1]/btn[9]").Press
    session.FindById("wnd[1]/tbar[0]/btn[8]").Press
    'session.findById("wnd[2]/tbar[0]/btn[0]").Press
    'session.findById("wnd[0]/usr/ctxtSO_MATNR-LOW").caretPosition = 8
Else
    'session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = codes
    session.FindById("wnd[0]/usr/ctxtSO_MATNR-LOW").Text = codes
    session.FindById("wnd[0]/usr/ctxtSO_MATNR-LOW").CaretPosition = 8
    session.FindById("wnd[0]/usr/ctxtSO_BEDAT-HIGH").SetFocus
    session.FindById("wnd[0]/usr/ctxtSO_BEDAT-HIGH").CaretPosition = 10
    session.FindById("wnd[0]/tbar[1]/btn[8]").Press
End If
'session.FindById("wnd[0]/usr/ctxtSO_MATNR-LOW").CaretPosition = 8

'PO Tab - Select All and Adopt
'session.FindById("wnd[0]/usr/ctxtSO_BEDAT-HIGH").SetFocus
'session.FindById("wnd[0]/usr/ctxtSO_BEDAT-HIGH").CaretPosition = 10
'session.FindById("wnd[0]/tbar[1]/btn[8]").Press
'session.FindById("wnd[0]/tbar[1]/btn[5]").Press
'session.FindById("wnd[0]/tbar[1]/btn[9]").Press
'session.FindById("wnd[1]/tbar[0]/btn[8]").Press
'session.FindById("wnd[2]/tbar[0]/btn[0]").Press

  


'session.FindById("wnd[0]/usr/ctxtSO_MATNR-LOW").CaretPosition = 8

'PO Tab - Select All and Adopt
'session.FindById("wnd[0]/usr/ctxtSO_BEDAT-HIGH").SetFocus
'session.FindById("wnd[0]/usr/ctxtSO_BEDAT-HIGH").CaretPosition = 10
'session.FindById("wnd[0]/tbar[1]/btn[8]").Press
session.FindById("wnd[0]/tbar[1]/btn[5]").Press
session.FindById("wnd[0]/tbar[1]/btn[9]").Press
'session.FindById("wnd[1]/tbar[0]/btn[8]").Press
'session.FindById("wnd[2]/tbar[0]/btn[0]").Press

'Payment Tab

PaymentTab = "wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_PAY/"
session.FindById("wnd[0]").SendVKey 0


session.FindById(PaymentTab).Select
session.FindById(PaymentTab & "ssubHEADER_SCREEN:SAPLFDCB:0020/ctxtINVFO-ZFBDT").Text = wrksht.Cells(row, lookfor(wrksht, "BaselineDate"))
session.FindById(PaymentTab & "ssubHEADER_SCREEN:SAPLFDCB:0020/ctxtINVFO-ZTERM").Text = "Z005"
session.FindById("wnd[0]").SendVKey 0
session.FindById("wnd[0]").SendVKey 0

'Go to Basic Data
session.FindById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_TOTAL").Select

            'GL FI
'Please add the GL account for all the FI invoice - Rebecca using the IF function


'loop through each row
  
      
'Payment Tab
'session.FindById(PaymentTab).Select
'session.FindById(PaymentTab & "ssubHEADER_SCREEN:SAPLFDCB:0020/ctxtINVFO-ZFBDT").Text = wrksht.Cells(row, lookfor(wrksht, "BaselineDate"))
'session.FindById(PaymentTab & "ssubHEADER_SCREEN:SAPLFDCB:0020/ctxtINVFO-ZTERM").Text = "z005"
'session.FindById("wnd[0]").SendVKey 0
'session.FindById("wnd[0]").SendVKey 0

'Go to Basic Data
session.FindById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_TOTAL").Select

            'GL FI
'Please add the GL account for all the FI invoice - Rebecca using the IF function



 
      session.FindById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/subITEMS:SAPLMR1M:6010/tabsITEMTAB/tabpITEMS_G/L").Select
      saprowIndex = 0
    

For j = 0 To lastRow - 2
    valuecolumnc = wrksht.Cells(j + 3, "I").Value
    If valuecolumnc < 0 Then
        keys = "S"
    Else
        keys = "H"
    End If

    session.FindById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/subITEMS:SAPLMR1M:6010/tabsITEMTAB/tabpITEMS_G/L/ssubTABS:SAPLMR1M:6040/ssubSACHKONTO:SAPLFSKB:0100/tblSAPLFSKBTABLE/ctxtACGL_ITEM-HKONT[1," & j & "]").Text = "520001006"
    session.FindById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/subITEMS:SAPLMR1M:6010/tabsITEMTAB/tabpITEMS_G/L/ssubTABS:SAPLMR1M:6040/ssubSACHKONTO:SAPLFSKB:0100/tblSAPLFSKBTABLE/txtACGL_ITEM-WRBTR[4," & saprowIndex & "]").Text = Abs(valuecolumnc)
    session.FindById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/subITEMS:SAPLMR1M:6010/tabsITEMTAB/tabpITEMS_G/L/ssubTABS:SAPLMR1M:6040/ssubSACHKONTO:SAPLFSKB:0100/tblSAPLFSKBTABLE/cmbACGL_ITEM-SHKZG[3," & j & "]").key = keys
    session.FindById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/subITEMS:SAPLMR1M:6010/tabsITEMTAB/tabpITEMS_G/L/ssubTABS:SAPLMR1M:6040/ssubSACHKONTO:SAPLFSKB:0100/tblSAPLFSKBTABLE/txtACGL_ITEM-WRBTR[4," & j & "]").SetFocus
    session.FindById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/subITEMS:SAPLMR1M:6010/tabsITEMTAB/tabpITEMS_G/L/ssubTABS:SAPLMR1M:6040/ssubSACHKONTO:SAPLFSKB:0100/tblSAPLFSKBTABLE/txtACGL_ITEM-WRBTR[4," & j & "]").CaretPosition = 9
    session.FindById("wnd[0]").SendVKey 0
    session.FindById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/subITEMS:SAPLMR1M:6010/tabsITEMTAB/tabpITEMS_G/L/ssubTABS:SAPLMR1M:6040/ssubSACHKONTO:SAPLFSKB:0100/tblSAPLFSKBTABLE/ctxtACGL_ITEM-GSBER[12," & j & "]").Text = "7400"
    session.FindById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/subITEMS:SAPLMR1M:6010/tabsITEMTAB/tabpITEMS_G/L/ssubTABS:SAPLMR1M:6040/ssubSACHKONTO:SAPLFSKB:0100/tblSAPLFSKBTABLE/ctxtACGL_ITEM-GSBER[12," & j & "]").CaretPosition = 4
    session.FindById("wnd[0]").SendVKey 0
    saprowIndex = saprowIndex + 1
Next j
'Exit Sub
response = MsgBox("Is everything correct in SAP?", vbQuestion + vbYesNo, "Verification")
If response = vbYes Then
    MsgBox "please perform necessary manual action in SAP and click 'OK' when ready to proceed.", vbInformation, "Manual Step"
    ' wait for user to click ok
Else
      Exit Function
End If
  
'Nextrow:
'row = row + 1
Wend
  
End Function
 
Public Function GetCodes(ByVal materialplantcodes As Object, ByVal key As String) As Variant
    Dim codes As Variant
    If materialplantcodes.Exists(key) Then
        GetCodes = materialplantcodes(key)
    Else
    GetCodes = "unknown"
    End If
End Function


'Enter Key
'session.FindById("wnd[0]").SendVKey 0

'Exit Sub
'MsgBox ("Please VERIFY all the data and click on POST manually. If there is an issue, click on the 'X' to exit the macro")

'Raisa please add the if function, if the invoice is not posted for indexing purposes.

'Next row


