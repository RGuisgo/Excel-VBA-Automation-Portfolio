VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} WorkbookSelector 
   Caption         =   "Workbook selector"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4605
   OleObjectBlob   =   "WorkbookSelector.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "WorkbookSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_Browse_Click()
On Error Resume Next

    Me.TextBox_HistoricalFile = Application.GetOpenFilename

End Sub

Private Sub Cmd_Cancel_Click()
On Error Resume Next

    Unload WorkbookSelector

End Sub

Private Sub cmd_format_Click()
On Error Resume Next

'Checking if all necessary data is given
    If ComboBox_Workbook.Value = "" Then
        MsgBox "Please select a Workbook"
        ComboBox_Workbook.SetFocus
        Exit Sub
    ElseIf ComboBox_Worksheet.Value = "" Then
        MsgBox "Please select a Worksheet"
        ComboBox_Worksheet.SetFocus
        Exit Sub
    ElseIf ComboBox_TypeOfMacro.Value = "" Then
        MsgBox "Please select Type of Macro to run"
        ComboBox_TypeOfMacro.SetFocus
        Exit Sub
    ElseIf TextBox_HistoricalFile.Value = "" Then
        MsgBox "Please select last day's file"
        TextBox_HistoricalFile.SetFocus
        Exit Sub
    ElseIf ComboBox_TypeOfMacro.Value = "Price Change Macro" Then
        PriceChangeMacro
        Exit Sub
    ElseIf ComboBox_TypeOfMacro.Value = "NomKey Macro" Then
        NomKeyMacro
        Exit Sub
    ElseIf ComboBox_TypeOfMacro.Value = "BCC Macro" Then
        BCCMacro
        Exit Sub
    End If
    
End Sub

Private Sub ComboBox_Workbook_Click()
On Error Resume Next

    Dim WorkbookSelected As String
    Dim ws As Worksheet

'Activate selected workbook
    WorkbookSelected = ComboBox_Workbook.Value
    Workbooks(WorkbookSelected).Activate
    
'Set ComboBox_WorkSheet Values
        ComboBox_Worksheet.Clear
    For Each ws In Worksheets
        ComboBox_Worksheet.AddItem ws.Name
    Next

End Sub

Private Sub ComboBox_Worksheet_Click()
On Error Resume Next

    Dim WorksheetSelected As String

'Activate selected worksheet
    WorksheetSelected = ComboBox_Worksheet.Value
    Worksheets(WorksheetSelected).Activate

End Sub

Private Sub Label_HistoricalFile_Click()

End Sub

Private Sub Label1_Click()

End Sub

Private Sub Label2_Click()

End Sub

Private Sub UserForm_Activate()
On Error Resume Next

'Choose type of macro
    With ComboBox_TypeOfMacro
        .AddItem "Price Change Macro"
        .AddItem "NomKey Macro"
        .AddItem "BCC Macro"
    End With

'Set ComboBox_WorkBook Values
    Dim wb As Workbook

    For Each wb In Workbooks
        ComboBox_Workbook.AddItem wb.Name
    Next

ComboBox_Workbook.SetFocus

End Sub

