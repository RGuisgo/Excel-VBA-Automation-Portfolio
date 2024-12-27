Attribute VB_Name = "NomKey"
Sub NomKeyMacro()
    On Error GoTo ErrorHandler ' Add error handling at the start
    
    Debug.Print "Subroutine started" ' Simple debug message to confirm the subroutine is running
    
    Dim i As Long
    Dim j As Long
    Dim WorkbookSelected As String
    Dim WorkbookToCompare As Workbook
    Dim lastRow1 As Long
    Dim lastRows As Long
    Dim settingsValue As Variant
    Dim filteredValue As Variant
    Dim filePath As String
    
    ' Turn off screen updating and events
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    ' Initialize variables
    WorkbookSelected = WorkbookSelector.ComboBox_Workbook.Value ' Replace with your actual workbook name
    If WorkbookSelected = "" Then
        MsgBox "No file selected. Exiting subroutine."
        GoTo Cleanup
    End If
    ' Debugging output
    'Debug.Print "WorkbookSelected: " & WorkbookSelected
    'Debug.Print "TextBox_HistoricalFile value: " & WorkbookSelector.TextBox_HistoricalFile.Value
    
    ' Open the selected workbook
    filePath = WorkbookSelector.TextBox_HistoricalFile.Value
    If filePath = "" Then
        MsgBox "No file selected. Exiting subroutine."
        GoTo Cleanup
    End If
    
    Set WorkbookToCompare = Workbooks.Open(filePath)
    
    If WorkbookToCompare Is Nothing Then
        MsgBox "The workbook specified in TextBox_HistoricalFile could not be found. Exiting subroutine."
        GoTo Cleanup
    End If
    
    ' Debugging output
    'Debug.Print "WorkbookToCompare: " & WorkbookToCompare.Name
    
    ' Select the first sheet
    With WorkbookToCompare.Sheets(1)
        ' Find the last row in the new workbook's sheet
        lastRows = .Range("B" & .Rows.Count).End(xlUp).Row

        ' Debugging output
        'Debug.Print "Last row in new workbook's sheet: " & lastRows

        ' Find the last row in the Filtered worksheet
        lastRow1 = Workbooks(WorkbookSelected).Worksheets(1).Range("B" & Workbooks(WorkbookSelected).Worksheets(1).Rows.Count).End(xlUp).Row

        ' Debugging output
        'Debug.Print "Last row in Filtered: " & lastRow1

        ' Loop through the rows in the new workbook's sheet
        For i = 3 To lastRows
            ' Debugging output
            'Debug.Print "Processing new workbook's sheet row: " & i

            ' Loop through the rows in the Filtered worksheet
            For j = 2 To lastRow1
                ' Get values to compare
                settingsValue = .Range("B" & i).Value
                filteredValue = Workbooks(WorkbookSelected).Worksheets(1).Range("B" & j).Value

                ' Debugging output
                'Debug.Print "Comparing new workbook's sheet B" & i & " (" & settingsValue & ") with Filtered!B" & j & " (" & filteredValue & ")"

                ' Compare values
                If settingsValue = filteredValue Then
                    If IsEmpty(Workbooks(WorkbookSelected).Worksheets(1).Range("AF" & j)) Then
                    ' Copy value
                        Workbooks(WorkbookSelected).Worksheets(1).Range("AF" & j).Value = .Range("I" & i).Value
                    'Debug.Print "Value copied from new workbook's sheet I" & i & " to Filtered!AF" & j
                    End If
                End If
            Next j
        Next i
    End With
    
     ' Autofit columns in the selected workbook
    Workbooks(WorkbookSelected).Worksheets(1).Columns.AutoFit



    'Debug.Print "Completed CompareAndCopy subroutine"

Cleanup:
    ' Turn on screen updating and events
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic

    ' Close the macro
    Application.StatusBar = "*** Done ***"
    If Not WorkbookToCompare Is Nothing Then WorkbookToCompare.Close False
    Workbooks(WorkbookSelected).Activate
    Unload WorkbookSelector
    DoEvents ' Allow Excel to process any pending events
    Application.StatusBar = False

    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description
    Resume Cleanup

End Sub

