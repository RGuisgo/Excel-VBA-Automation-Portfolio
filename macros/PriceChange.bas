Attribute VB_Name = "PriceChange"
Sub PriceChangeMacro()
On Error Resume Next

    Dim WorkbookSelected, SheetSelected, HistoricalFileSelected, HistoricalFile, WorkbookNew As String
    Dim PrefixToUse, LenPrefixToUse, settingsValue, filteredValue As Variant
    Dim HistoricalDocNo, NewDocNo, HistoricalCC, HistoricalStatus, NewStatus, TNToFilter, HistoricalTNo, NewCC As Long
    Dim DocTypeReceiverToFilter, VendorToFilter, DocTypeOriginatorToFilter As String
    Dim InvNrToFilter, DCToFilter As String
    Dim lastRow, lastRows, lastRowf, lastRow1, LastRowHistorical, i, j As Long

 'Deleting unnecessary columns
 
    Range("M:M").Select
    Selection.Delete Shift:=xlToLeft
  

'Turning off calculation and screenupdating to speed up formatting
    Application.StatusBar = "Turning off calculation and screenupdating to speed up formatting"
    With Application
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
        .DisplayStatusBar = True
    End With

'Setting default values
    Application.StatusBar = "Setting default values"
        WorkbookSelected = WorkbookSelector.ComboBox_Workbook.Value
        SheetSelected = WorkbookSelector.ComboBox_Worksheet.Value
        HistoricalFileSelected = WorkbookSelector.TextBox_HistoricalFile.Value

'Copying sheet
    Application.StatusBar = "Copying sheet"
    
        ActiveSheet.Copy Before:=Worksheets(SheetSelected)
        ActiveSheet.Name = "Filtered"

'Deleting empty and unnecessary rows & columns
    Application.StatusBar = "Deleting empty and unnecessary rows & columns"
    
        'Deleting unnecessary columns
'        Range("B:C,E:E,G:G,I:I,L:M,O:P,R:R,T:AJ,AM:AO").Select
        Range("A:B,E:E,G:G,I:I").Select
            Selection.Delete Shift:=xlToLeft
    
          Columns("X:X").Select
          Selection.Insert Shift:=xlToRight

        'Deleting empty rows
        If WorksheetFunction.CountA(Cells) > 0 Then
            lastRow = Cells.Find(What:="*", After:=[A1], _
                  SearchOrder:=xlByRows, _
                  SearchDirection:=xlPrevious).Row
        End If
        Range("A1" & ":" & "AG" & CLng(lastRow)).Select
        For i = Selection.Rows.Count To 1 Step -1
            If WorksheetFunction.CountA(Selection.Rows(i)) = 0 Then
                Selection.Rows(i).EntireRow.Delete
            End If
            Application.StatusBar = "Deleting empty rows " & i
        Next i

        'Deleting unnecessary rows
        If WorksheetFunction.CountA(Cells) > 0 Then
            lastRow = Cells.Find(What:="*", After:=[A1], _
                  SearchOrder:=xlByRows, _
                  SearchDirection:=xlPrevious).Row
        End If
        Range("A1" & ":" & "AG" & CLng(lastRow)).Select
        For i = Selection.Rows.Count To 1 Step -1
            If Range("E" & i).Value = "" And Range("B" & i).Value = "" Then
                Selection.Rows(i).EntireRow.Delete
            End If
            Application.StatusBar = "Deleting unnecessary rows " & i
        Next i
        
       
    
        'Deleting emtpy rows
        If WorksheetFunction.CountA(Cells) > 0 Then
            lastRow = Cells.Find(What:="*", After:=[A1], _
                  SearchOrder:=xlByRows, _
                  SearchDirection:=xlPrevious).Row
        End If
        Range("A1" & ":" & "A" & CLng(lastRow)).Select
        For i = Selection.Rows.Count To 1 Step -1
            If Range("A" & i).Value = "" Then
                Selection.Rows(i).EntireRow.Delete
            End If
            Application.StatusBar = "Deleting emtpy rows " & i
        Next i

'Formatting header
    Application.StatusBar = "Formatting header"
       Rows("1:1").Select
            Selection.Insert Shift:=xlDown
        Range("A1").FormulaR1C1 = "CC"
        Range("B1").FormulaR1C1 = "Trade Num"
        Range("C1").FormulaR1C1 = "Item"
        Range("D1").FormulaR1C1 = "Material"
        Range("E1").FormulaR1C1 = "Pur. Doc."
        Range("F1").FormulaR1C1 = "Item"
        Range("G1").FormulaR1C1 = "Nom. Key"
        Range("H1").FormulaR1C1 = "Item"
    '   Range("I1").FormulaR1C1 = "Doc. No."
        Range("J1").FormulaR1C1 = "Doc. No."
        Range("K1").FormulaR1C1 = "Year"
      ' Range("L1").FormulaR1C1 = "Created On"
        Range("M1").FormulaR1C1 = "Item"
        Range("N1").FormulaR1C1 = "Created On"
        Range("O1").FormulaR1C1 = "Invoice date"
        Range("P1").FormulaR1C1 = "Formula"
        Range("Q1").FormulaR1C1 = "Doc. Amt."
        Range("R1").FormulaR1C1 = "Crcy"
        Range("S1").FormulaR1C1 = "UoM"
        Range("T1").FormulaR1C1 = "New Amt."
        Range("U1").FormulaR1C1 = "Crcy"
        Range("V1").FormulaR1C1 = "UoM"
        Range("W1").FormulaR1C1 = "Tot. Doc. Amt."
        Range("X1").FormulaR1C1 = "Tot. New Amt."
        Range("Y1").FormulaR1C1 = "Difference Amt."
        Range("Z1").FormulaR1C1 = "Abs. Difference Amt."
        Range("AA1").FormulaR1C1 = "Crcy"
        Range("AB1").FormulaR1C1 = "MT"
        Range("AC1").FormulaR1C1 = "Material Description"
        Range("AD1").FormulaR1C1 = "Vessel Name"
        Range("AE1").FormulaR1C1 = "Short Description"
        Range("AF1").FormulaR1C1 = "Status"
        Range("AG1").FormulaR1C1 = "Short Description"
        'Range("AH1").FormulaR1C1 = "Vendor Name"
        

                
                
          
          'Handling historical and new items
          
          
          
            Application.StatusBar = "Handling historical and new items"
                
                If WorksheetFunction.CountA(Cells) > 0 Then
                    lastRow = Cells.Find(What:="*", After:=[A1], _
                          SearchOrder:=xlByRows, _
                          SearchDirection:=xlPrevious).Row
                End If
                
                If WorkbookSelector.CheckBox_Save.Value = True Then
                    Workbooks.Add
                    Sheets("Sheet1").Name = "Cleared"
                    WorkbookNew = ActiveWorkbook.Name
                End If
                
                Workbooks.Open Filename:=HistoricalFileSelected
                HistoricalFile = ActiveWorkbook.Name
                
                If WorksheetFunction.CountA(Cells) > 0 Then
                    LastRowHistorical = Cells.Find(What:="*", After:=[A1], _
                          SearchOrder:=xlByRows, _
                          SearchDirection:=xlPrevious).Row
                End If
                
                ' To unfreeze historical file worksheet
                Range("A1").Select
                ActiveWindow.FreezePanes = True
                ActiveWindow.FreezePanes = False
                
                
                
                'Copying comments for historical items
                
                Range("A1" & ":" & "A" & CLng(LastRowHistorical)).Select
                For i = Selection.Rows.Count To 2 Step -1
                'For i = LastRowHistorical To 2 Step -1
                    HistoricalStatus = Cells(i, "AD").Value
                
                    If LCase(HistoricalStatus) = "pending" Then
                        Rows(i).Copy Destination:=Workbooks(WorkbookSelected).Worksheets("Filtered").Cells(Rows.Count, 1).End(xlUp).Offset(1, 0)
                        Workbooks(HistoricalFile).Activate
                        Cells(i, "AG").Value = "x"
                        GoTo NextIteration ' Skip the rest of the code for this iteration
                    End If
                
                    Application.StatusBar = "Copying comments for historical items " & i
NextIteration:
                Next i
                
                
                'Copying cleared items into a new file
                If WorkbookSelector.CheckBox_Save.Value = True Then
                    Range("A1" & ":" & "A" & CLng(LastRowHistorical)).Select
                    For i = Selection.Rows.Count To 2 Step -1
                        If Range("AG" & i).Value <> "x" Then
                            Range("A" & i & ":" & "AG" & i).Select
                                Selection.Copy
                            Workbooks(WorkbookNew).Worksheets("Cleared").Activate
                            Range("A" & i & ":" & "AG" & i).Select
                                ActiveSheet.Paste
                            Workbooks(HistoricalFile).Activate
                        End If
                        Application.StatusBar = "Copying cleared items into a new file " & i
                    Next i
                End If


                'Marking new items
                Workbooks(WorkbookSelected).Worksheets("Filtered").Activate
                Range("A1" & ":" & "A" & CLng(lastRow)).Select
                For i = Selection.Rows.Count To 2 Step -1
                    If WorksheetFunction.CountA(Range("AC" & i & ":" & "AD" & i)) = 0 Then
                        Range("A" & i & ":" & "AG" & i).Interior.ColorIndex = 6
                    End If
                    Application.StatusBar = "Marking new items " & i
                Next i
                
                 'Getting absolute values and number formatting
            
            Application.StatusBar = "Getting absolute values and number formatting"
            With Application
                '.Range("I1:I" & lastRow).NumberFormat = "General"
                .Range("X1:X" & lastRow).FormulaR1C1 = "=ABS(RC[-1])"
            End With
            Application.StatusBar = "Getting absolute values and number formatting"
               
                
                
        
                
     'Deleting  rows with absolute values difference in price less than 1000
     
                If WorksheetFunction.CountA(Cells) > 0 Then
                    lastRow = Cells.Find(What:="*", After:=[A1], _
                          SearchOrder:=xlByRows, _
                          SearchDirection:=xlPrevious).Row
                End If
                Range("A1" & ":" & "AG" & CLng(lastRow)).Select
                For i = Selection.Rows.Count To 2 Step -1
                    If Range("X" & i).Value <= 1000 Then
                        Selection.Rows(i).EntireRow.Delete
                    End If
                    Application.StatusBar = "Deleting unnecessary rows " & i
                Next i
               
               
   


        'Finalizing cleared items worksheet
            Application.StatusBar = "Finalizing cleared items worksheet"

                If WorkbookSelector.CheckBox_Save.Value = True Then
                    Workbooks(WorkbookNew).Activate
                    Range("A1").FormulaR1C1 = "CC"
                    Range("B1").FormulaR1C1 = "Trade Num"
                    Range("C1").FormulaR1C1 = "Item"
                    Range("D1").FormulaR1C1 = "Material"
                    Range("E1").FormulaR1C1 = "Pur. Doc."
                    Range("F1").FormulaR1C1 = "Item"
                    Range("G1").FormulaR1C1 = "Nom. Key"
                    Range("H1").FormulaR1C1 = "Item"
                    Range("I1").FormulaR1C1 = "Doc. No."
                    Range("J1").FormulaR1C1 = "Year"
                    Range("K1").FormulaR1C1 = "Item"
                    Range("L1").FormulaR1C1 = "Created On"
                    Range("M1").FormulaR1C1 = "Invoice date"
                    Range("N1").FormulaR1C1 = "Formula"
                    Range("O1").FormulaR1C1 = "Doc. Amt."
                    Range("P1").FormulaR1C1 = "Crcy"
                    Range("Q1").FormulaR1C1 = "UoM"
                    Range("R1").FormulaR1C1 = "New Amt."
                    Range("S1").FormulaR1C1 = "Crcy"
                    Range("T1").FormulaR1C1 = "UoM"
                    Range("U1").FormulaR1C1 = "Tot. Doc. Amt."
                    Range("V1").FormulaR1C1 = "Tot. New Amt."
                    Range("W1").FormulaR1C1 = "Difference Amt."
                    Range("X1").FormulaR1C1 = "Abs. Difference Amt."
                    Range("Y1").FormulaR1C1 = "Crcy"
                    Range("Z1").FormulaR1C1 = "MT"
                    Range("AA1").FormulaR1C1 = "Material Description"
                    Range("AB1").FormulaR1C1 = "Vessel Name"
                    Range("AC1").FormulaR1C1 = "Short Description"
                    Range("AD1").FormulaR1C1 = "Status"
                    Range("AE1").FormulaR1C1 = "Short Description"
                    Range("AF1").FormulaR1C1 = "Vendor Name"
                    Range("AG1").FormulaR1C1 = "Receiving Date"
                    
                    
               
                   


                    'Formatting header
                    Application.StatusBar = "Formatting header"
                    Range("A1:AG1").Select
                        Selection.Font.Bold = True
                    With Selection
                        .WrapText = True
                        .VerticalAlignment = xlCenter
                    End With

                    'Setting activecells back to A1
                    Application.StatusBar = "Setting activecells back to A1"
                    Range("A1").Select
                    ActiveWindow.ScrollRow = 1
                    ActiveWindow.ScrollColumn = 1

                    'AutoFitting columns width
                    Application.StatusBar = "AutoFitting columns width"
                    Cells.Select
                    Cells.EntireColumn.AutoFit
                    Range("A1").Select
                End If
                

'        'Formatting header
'            Application.StatusBar = "Formatting header"
                Range("A1").FormulaR1C1 = "CC"
                Range("B1").FormulaR1C1 = "Trade Num"
                Range("C1").FormulaR1C1 = "Item"
                Range("D1").FormulaR1C1 = "Material"
                Range("E1").FormulaR1C1 = "Pur. Doc."
                Range("F1").FormulaR1C1 = "Item"
                Range("G1").FormulaR1C1 = "Nom. Key"
                Range("H1").FormulaR1C1 = "Item"
                Range("I1").FormulaR1C1 = "Doc. No."
                Range("J1").FormulaR1C1 = "Year"
                Range("K1").FormulaR1C1 = "Item"
                Range("L1").FormulaR1C1 = "Created On"
                Range("M1").FormulaR1C1 = "Invoice date"
                Range("N1").FormulaR1C1 = "Formula"
                Range("O1").FormulaR1C1 = "Doc. Amt."
                Range("P1").FormulaR1C1 = "Crcy"
                Range("Q1").FormulaR1C1 = "UoM"
                Range("R1").FormulaR1C1 = "New Amt."
                Range("S1").FormulaR1C1 = "Crcy"
                Range("T1").FormulaR1C1 = "UoM"
                Range("U1").FormulaR1C1 = "Tot. Doc. Amt."
                Range("V1").FormulaR1C1 = "Tot. New Amt."
                Range("W1").FormulaR1C1 = "Difference Amt."
                Range("X1").FormulaR1C1 = "Abs. Difference Amt."
                Range("Y1").FormulaR1C1 = "Crcy"
                Range("Z1").FormulaR1C1 = "MT"
                Range("AA1").FormulaR1C1 = "Material Description"
                Range("AB1").FormulaR1C1 = "Vessel Name"
                Range("AC1").FormulaR1C1 = "Short Description"
                'Range("AC1").FormulaR1C1 = "Short Description"
                Range("AD1").FormulaR1C1 = "Status"
                Range("AE1").FormulaR1C1 = "Short Description"
                Range("AF1").FormulaR1C1 = "Vendor Name"
                Range("AG1").FormulaR1C1 = "Receiving Date"
                
                
              
                       'Setting clearing date
                Application.StatusBar = "Setting clearing date"
                If WorksheetFunction.CountA(Cells) > 0 Then
                    lastRow = Cells.Find(What:="*", After:=[A1], _
                          SearchOrder:=xlByRows, _
                          SearchDirection:=xlPrevious).Row
                End If
                Range("A1" & ":" & "A" & CLng(lastRow)).Select
                For i = Selection.Rows.Count To 2 Step -1
                    If IsEmpty(Cells(i, "AG")) Then
                    Range("AG" & i).Value = Format(Date, "mm/dd/yyyy")
                    End If
                Next i
                
                Application.StatusBar = "Clearing date set"

'
        'Closing Historical file
            Application.StatusBar = "Closing Historical file"
                Application.DisplayAlerts = False
                    Workbooks(HistoricalFile).Activate
                    ActiveWindow.Close
                Application.DisplayAlerts = True

        'Setting activecells back to A1
            Application.StatusBar = "Setting activecells back to A1"
                Workbooks("IC-TP PRICE Macro  Nov2024.xls").Activate
                    Range("A1").Select
                    ActiveWindow.ScrollRow = 1
                    ActiveWindow.ScrollColumn = 1
                    Worksheets("Macro").Activate
                Workbooks(WorkbookSelected).Activate
                    Worksheets(SheetSelected).Activate
                        Range("A1").Select
                        ActiveWindow.ScrollRow = 1
                        ActiveWindow.ScrollColumn = 1
                    Worksheets("Filtered").Activate
                        Range("A1").Select
                        ActiveWindow.ScrollRow = 1
                        ActiveWindow.ScrollColumn = 1

        'Naming tabs
            Application.StatusBar = "Naming tabs"
                Worksheets(SheetSelected).Name = "Raw_Data"
                Worksheets("Filtered").Name = "PRICE-" & Format(Date, "mmddyyyy")

        'Color formatting header
            Application.StatusBar = "Color formatting header"
                Columns("A:AG").Select
                With Selection
                    .HorizontalAlignment = xlCenter
                End With
                Range("A1:C1").Select
                    Selection.Interior.ColorIndex = 53
                Range("D1:J1").Select
                    Selection.Interior.ColorIndex = 45
                Range("K1").Select
                    Selection.Interior.ColorIndex = 53
                Range("L1").Select
                    Selection.Interior.ColorIndex = 45
                Range("M1").Select
                    Selection.Interior.ColorIndex = 53
                Range("N1").Select
                    Selection.Interior.ColorIndex = 43
                Range("O1:Q1").Select
                    Selection.Interior.ColorIndex = 45
                Range("R1:T1").Select
                    Selection.Interior.ColorIndex = 53
                Range("U1:Z1").Select
                    Selection.Interior.ColorIndex = 43
                Range("AA1:AD1").Select
                    Selection.Interior.ColorIndex = 45
                Range("AE1:AG1").Select
                    Selection.Interior.ColorIndex = 53
                Range("A1:AG1").Select
                    Selection.Font.Bold = True
                With Selection
                    .WrapText = True
                    .VerticalAlignment = xlCenter
                End With

        'AutoFitting columns width
            Application.StatusBar = "AutoFitting columns width"
                Cells.Select
                Cells.EntireColumn.AutoFit
                Range("A1").Select

        'Setting back calculation and screenupdating
            With Application
                .Calculation = xlCalculationAutomatic
                .DisplayStatusBar = True
            End With


        If WorksheetFunction.CountA(Cells) > 0 Then
                        lastRow = Cells.Find(What:="*", After:=[A1], _
                              SearchOrder:=xlByRows, _
                              SearchDirection:=xlPrevious).Row
        End If

'        For i = 2 To lastRow Step 1
'            Age = Format(Date, "mm") - Format(Range("D" & i).Value, "mm") + 12 * (Format(Date, "yy") - Format(Range("D" & i).Value, "yy"))
'
'            Select Case Age
'                Case "0"
'                    Range("O" & i).Value = " 1-30"
'                Case "1"
'                    Range("O" & i).Value = "30-60"
'                Case Else
'                    Range("O" & i).Value = "Aged"
'            End Select
'
'            Application.StatusBar = "Allocating aging category " & i
'
'        Next i
'

'
'
'
'            Range("P2:P" & lastRow).Select
'            With Selection.Validation
'                .Delete
'                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
'                xlBetween, Formula1:="=$AF$2:$AF$22"
'                .IgnoreBlank = True
'                .InCellDropdown = True
'                .InputTitle = ""
'                .ErrorTitle = ""
'                .InputMessage = ""
'                .ErrorMessage = ""
'                .ShowInput = True
'                .ShowError = True
'            End With
'
'            Columns("AF:AF").Select
'            Selection.EntireColumn.Hidden = True
'            Range("A1").Select
'
'            For i = 2 To lastRow
'            Range("P" & i).Copy
'            Range("A" & i & ":I" & i).Select
'            Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
'            Application.CutCopyMode = False
'
'            Next i
'

        ' Set up colums width and alignment
            Columns("A:A").ColumnWidth = 7
            Columns("B:B").ColumnWidth = 7
            Columns("C:C").ColumnWidth = 10.14
            Columns("D:D").ColumnWidth = 9.14
            Columns("E:E").ColumnWidth = 9.14
            Columns("F:F").ColumnWidth = 4.43
            Columns("G:G").ColumnWidth = 3.14
            Columns("H:H").ColumnWidth = 8.43
            Columns("I:I").ColumnWidth = 12.7
            Columns("I:I").Select
            Selection.NumberFormat = "General"
            Columns("J:J").ColumnWidth = 8.43
            Columns("K:K").ColumnWidth = 11.14
            Columns("L:L").ColumnWidth = 12.71
            Columns("M:M").ColumnWidth = 12.57
            Columns("N:N").ColumnWidth = 6.86
            Columns("O:O").ColumnWidth = 18.86
            Columns("P:P").ColumnWidth = 9.87
            Columns("Q:Q").ColumnWidth = 11.14
            Columns("R:R").ColumnWidth = 12.71
            Columns("S:S").ColumnWidth = 12.57
            Columns("T:T").ColumnWidth = 6.86
            Columns("U:U").ColumnWidth = 18.86
            Columns("V:V").ColumnWidth = 9.87
            Columns("W:W").ColumnWidth = 11.14
            Columns("X:Y").ColumnWidth = 12.71
            Columns("Z:Z").ColumnWidth = 12.57
            Columns("AA:AA").ColumnWidth = 6.86
            Columns("AB:AB").ColumnWidth = 6.86
            Columns("AC:AC").ColumnWidth = 71
            Columns("AC:AC").Select
            With Selection
                .HorizontalAlignment = xlLeft
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = False
            End With
            Columns("AD:AG").ColumnWidth = 9.14
            


'
'
'            ' Set date format
'            Columns("D:D").Select
'            Selection.NumberFormat = "m/d/yyyy"
'
'

            ' Freeze panels
            Rows("2:2").Select
            ActiveWindow.FreezePanes = True


            'Close the macro
            Application.StatusBar = "*** Done ***"

            'Workbooks("IC-TP PRICE Macro  Nov2024.xls").Close
            Workbooks(WorkbookSelected).Activate
            Unload WorkbookSelector
            Application.ScreenUpdating = True
            Application.StatusBar = False



            ' Set autofilter
            Rows("1:1").Select
            Selection.AutoFilter
End Sub


