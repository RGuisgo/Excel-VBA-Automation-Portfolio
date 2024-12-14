Attribute VB_Name = "mod_ChangeDocType"
Public Function ChangeDocType()

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

'Changing Doc Type
session.FindById("wnd[0]/usr/subSUB_MAIN:/OPT/SAPLVIM_IDX_UI:1001/subSUB_PROC_OPTIONS:/OPT/SAPLVIM_IDX_UI:1003/cntlCC_PROCESS_OPTIONS/shellcont/shell").CurrentCellRow = 2
session.FindById("wnd[0]/usr/subSUB_MAIN:/OPT/SAPLVIM_IDX_UI:1001/subSUB_PROC_OPTIONS:/OPT/SAPLVIM_IDX_UI:1003/cntlCC_PROCESS_OPTIONS/shellcont/shell").PressButtonCurrentCell
session.FindById("wnd[1]/usr/btnBUTTON_1").Press
session.FindById("wnd[1]/usr/cmbG_NEW_DOC_TYPE").key = "NPO_PMI_DE"
session.FindById("wnd[1]/tbar[0]/btn[5]").Press

'Next row
Nextrow:
row = row + 1

Wend

End Function
