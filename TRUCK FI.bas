Attribute VB_Name = "mod_FI"
Public Function FI()

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

Set wrkbk = Excel.Application.Workbooks(ThisWorkbook.Name)
Set wrksht = wrkbk.Sheets("FI")

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

                                                'FI Invoice - Done

 'Basic Data Tab
IndexData = "wnd[0]/usr/subSUB_MAIN:/OPT/SAPLVIM_IDX_UI:1001/subSUB_TAB_STRIP:/OPT/SAPLVIM_IDX_UI:1002/tabsTAB_MAIN/"
BasicData = IndexData & "tabpTAB1/ssubTAB_MAIN_SUBSCREEN:SAPLZF_VIM_IDX_UI:8001/ctxtGH_IDX_APPLICATION->MS_IDX_HEADER-"

'Do not auto post
session.FindById("wnd[0]/usr/subSUB_MAIN:/OPT/SAPLVIM_IDX_UI:1001/subSUB_TAB_STRIP:/OPT/SAPLVIM_IDX_UI:1002/tabsTAB_MAIN/tabpTAB1/ssubTAB_MAIN_SUBSCREEN:SAPLZF_VIM_IDX_UI:8001/chkGH_IDX_APPLICATION->MS_IDX_HEADER-CUSTOM_FIELD4").Selected = False

'Baseline Date
session.FindById(BasicData & "ZZBASELINE_DATE").Text = wrksht.Cells(row, lookfor(wrksht, "BaselineDate"))

'Payment Term
session.FindById(BasicData & "PYMNT_TERMS").Text = "Z005"

'Tax Code
session.FindById(BasicData & "TAX_CODE").Text = "VL"

'Basic Text
session.FindById(IndexData & "tabpTAB1/ssubTAB_MAIN_SUBSCREEN:SAPLZF_VIM_IDX_UI:8001/txtGH_IDX_APPLICATION->MS_IDX_HEADER-SGTXT").Text = wrksht.Cells(row, lookfor(wrksht, "Text"))

'Untick Auto Calc Tax Flag
session.FindById("wnd[0]/usr/subSUB_MAIN:/OPT/SAPLVIM_IDX_UI:1001/subSUB_TAB_STRIP:/OPT/SAPLVIM_IDX_UI:1002/tabsTAB_MAIN/tabpTAB1/ssubTAB_MAIN_SUBSCREEN:SAPLZF_VIM_IDX_UI:8001/chkGH_IDX_APPLICATION->MS_IDX_HEADER-AUTO_CALC").Selected = False

    'Vendor Info Tab
session.FindById(IndexData & "tabpTAB4").Select
VendorInfo = IndexData & "tabpTAB4/ssubTAB_MAIN_SUBSCREEN:SAPLZF_VIM_IDX_UI:8003/ctxtGH_IDX_APPLICATION->MS_IDX_HEADER-"

'Permitted Payee Number
session.FindById(VendorInfo & "ATTRIBUTE1").Text = "0000555446"

'Bank Account
session.FindById(VendorInfo & "BVTYP").Text = "0002"

    'Auto Post
On Error Resume Next
session.FindById("wnd[0]/usr/subSUB_MAIN:/OPT/SAPLVIM_IDX_UI:1001/subSUB_PROC_OPTIONS:/OPT/SAPLVIM_IDX_UI:1003/cntlCC_PROCESS_OPTIONS/shellcont/shell").CurrentCellRow = 3
session.FindById("wnd[0]/usr/subSUB_MAIN:/OPT/SAPLVIM_IDX_UI:1001/subSUB_PROC_OPTIONS:/OPT/SAPLVIM_IDX_UI:1003/cntlCC_PROCESS_OPTIONS/shellcont/shell").PressButtonCurrentCell
session.FindById("wnd[1]/usr/btnBUTTON_2").Press
On Error GoTo 0

'FI - GL Account
session.FindById("wnd[0]/usr/subITEMS:SAPLFSKB:0100/tblSAPLFSKBTABLE/ctxtACGL_ITEM-HKONT[1,0]").Text = "520001006"

session.FindById("wnd[0]").SendVKey 0
session.FindById("wnd[0]/usr/subITEMS:SAPLFSKB:0100/tblSAPLFSKBTABLE/cmbACGL_ITEM-SHKZG[3,0]").key = "S"
session.FindById("wnd[0]/usr/subITEMS:SAPLFSKB:0100/tblSAPLFSKBTABLE/txtACGL_ITEM-WRBTR[4,0]").Text = wrksht.Cells(row, lookfor(wrksht, "NetAmount"))
session.FindById("wnd[0]/usr/subITEMS:SAPLFSKB:0100/tblSAPLFSKBTABLE/ctxtACGL_ITEM-GSBER[15,0]").Text = "7400"

response = MsgBox("Is everything correct in SAP?", vbQuestion + vbYesNo, "Verification")
If response = vbYes Then
    MsgBox "please perform necessary manual action in SAP and click 'OK' when ready to proceed.", vbInformation, "Manual Step"
    ' wait for user to click ok
Else
      Exit Function
End If

'Enter - session.FindById("wnd[0]").SendVKey 0

'Next row
Nextrow:
row = row + 1
Wend

End Function

