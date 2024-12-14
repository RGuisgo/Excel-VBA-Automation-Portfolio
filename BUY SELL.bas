Attribute VB_Name = "BS"
Sub BuysellContractMacro()
On Error Resume Next

    Dim wsOriginal, wsMerged, wsSheet2, wsSheet3 As Worksheet
    Dim copyColumns1, copyColumns2, col, colsToCopy As Variant
    Dim pasteRange1, pasteRange2, pasteRange, copyRange, combinedRange1, combinedRange2, dataRange As Range
    Dim lastRow, lastRow1, lastRow2, lastRowG, lastRowMerged, i As Long
    Dim selectedWorkbook As String
    Dim sourceWookbook As Workbook
    Dim sourceWorkbookPath As String
    
      
      'Created by Raisa Lubis and Rebecca Guisgo 23-01-2024
      
    'Open the selected workbook
    'Prompt to select workbook
    
    sourceWorkbookPath = Application.GetOpenFilename(FileFilter:="Excel Files (*.xls; *.xlsx),*.xls; *xlsx", Title:="Select Workbook")
    
    'check workbook is selected
    
    If sourceWorkbook <> "False" Then
        'open selected workbook
        Set sourceWorkbook = Workbooks.Open(sourceWorkbookPath)
        
     'Set the original worksheet
    
        Set wsOriginal = sourceWorkbook.Sheets("Final")
    
    
    'deleting first row
    
     wsOriginal.Rows("1:1").Delete Shift:=xlUp
    
    'Deleting empty rows
        If WorksheetFunction.CountA(Cells) > 0 Then
            lastRow = Cells.Find(What:="*", After:=[A1], _
                  SearchOrder:=xlByRows, _
                  SearchDirection:=xlPrevious).Row
        End If
        Range("A1" & ":" & "U" & CLng(lastRow)).Select
        For i = Selection.Rows.Count To 1 Step -1
            If WorksheetFunction.CountA(Selection.Rows(i)) = 0 Then
                Selection.Rows(i).EntireRow.Delete
            End If
            Application.StatusBar = "Deleting empty rows " & i
        Next i
        
    
  
       'Set columns for the first set
    
        copyColumns1 = Array("C", "H", "I", "J", "K", "N", "U", "T")
    
    'Set columns for the second set
    
        copyColumns2 = Array("C", "H", "I", "J", "K", "O", "U", "T")
        
         'Set columns for the first set
    
        copyColumns3 = Array("C", "H", "I", "J", "K", "P", "U", "T")
    
    'Set columns for the second set
    
        copyColumns4 = Array("C", "H", "I", "J", "K", "Q", "U", "T")
        
      'create new worksheet for merged data
    
        Set wsMerged = sourceWorkbook.Sheets.Add(After:=sourceWorkbook.Sheets(sourceWorkbook.Sheets.Count))
        wsMerged.Name = "MergedData"
    
    
        
                'find last row with data in original sheet
    
        lastRow = wsOriginal.Cells(wsMerged.Rows.Count, "A").End(x1Up).Row
    


    
    'Set paste range in new sheet
        
        Set pasteRange = wsMerged.Range("A1")
    
    
    'create combined ranges
    
    
        For Each col In copyColumns1
            If combinedRange1 Is Nothing Then
                Set combinedRange1 = wsOriginal.Range(col & "1:" & col & lastRow)
            Else
                Set combinedRange1 = Union(combinedRange1, wsOriginal.Range(col & "1:" & col & lastRow))
            End If
        Next col
        
        
        
        For Each col In copyColumns2
            If combinedRange2 Is Nothing Then
                Set combinedRange2 = wsOriginal.Range(col & "1:" & col & lastRow)
            Else
                Set combinedRange2 = Union(combinedRange2, wsOriginal.Range(col & "1:" & col & lastRow))
            End If
        Next col
        
        For Each col In copyColumns3
            If combinedRange3 Is Nothing Then
                Set combinedRange3 = wsOriginal.Range(col & "1:" & col & lastRow)
            Else
                Set combinedRange3 = Union(combinedRange3, wsOriginal.Range(col & "1:" & col & lastRow))
            End If
        Next col
        
        
        
        For Each col In copyColumns4
            If combinedRange4 Is Nothing Then
                Set combinedRange4 = wsOriginal.Range(col & "1:" & col & lastRow)
            Else
                Set combinedRange4 = Union(combinedRange4, wsOriginal.Range(col & "1:" & col & lastRow))
            End If
        Next col

    
    'copy and paste the first set of columns
    
        combinedRange1.Copy Destination:=pasteRange
    
    'reset paste range for the second set of columns
    
    'Set pasteRange = pasteRange.Offeset(1, 0)
    
        nextEmptyRow = wsMerged.Cells(wsMerged.Rows.Count, 1).End(xlUp).Row + 1
     'copy and paste the first set of columns
    
        combinedRange2.Copy Destination:=wsMerged.Range("A" & nextEmptyRow)
        
        nextEmptyRow1 = wsMerged.Cells(wsMerged.Rows.Count, 1).End(xlUp).Row + 2
     'copy and paste the first set of columns
        combinedRange3.Copy Destination:=wsMerged.Range("A" & nextEmptyRow1)
        
        nextEmptyRow2 = wsMerged.Cells(wsMerged.Rows.Count, 1).End(xlUp).Row + 3
        
        combinedRange4.Copy Destination:=wsMerged.Range("A" & nextEmptyRow2)
    
        'Deleting unnecessary rows
        If WorksheetFunction.CountA(Cells) > 0 Then
            lastRow = Cells.Find(What:="*", After:=[A1], _
                  SearchOrder:=xlByRows, _
                  SearchDirection:=xlPrevious).Row
        End If
        Range("A1" & ":" & "H" & CLng(lastRow)).Select
        For i = Selection.Rows.Count To 2 Step -1
            If Range("F" & i).Value = "" Then
                Selection.Rows(i).EntireRow.Delete
            End If
            Application.StatusBar = "Deleting unnecessary rows " & i
        Next i
        
         'Deleting unnecessary rows
        If WorksheetFunction.CountA(Cells) > 0 Then
            lastRow = Cells.Find(What:="*", After:=[A1], _
                  SearchOrder:=xlByRows, _
                  SearchDirection:=xlPrevious).Row
        End If
        Range("A1" & ":" & "H" & CLng(lastRow)).Select
        For i = Selection.Rows.Count To 2 Step -1
            If Range("F" & i).Value = "Sale document (affiliate)" Or Range("F" & i).Value = "Sale document (Trading company)" Or Range("F" & i).Value = "Purchase document (affiliate)" Or Range("F" & i).Value = "Purchase document (Trading company)" Then
                Selection.Rows(i).EntireRow.Delete
            End If
            Application.StatusBar = "Deleting unnecessary rows " & i
        Next i
        
       
        
        ' find last row i column F
        
        lastRowG = wsMerged.Cells(wsMerged.Rows.Count, "G").End(xlUp).Row
        
        'loop through each row in column F
        
        For i = 2 To lastRowG
        
        'check the first three characters
        
            If Left(wsMerged.Cells(i, "F").Value, 3) = "400" Then
                wsMerged.Cells(i, "G").Value = "Sales"
            ElseIf Left(wsMerged.Cells(i, "F").Value, 3) = "470" Then
                wsMerged.Cells(i, "G").Value = "Purchase"
            End If
        Next i
        
       
            
 'Formatting header
    Application.StatusBar = "Formatting header"
        
        Range("A1").FormulaR1C1 = "Transaction Type"
        Range("B1").FormulaR1C1 = "Trading Company"
        Range("C1").FormulaR1C1 = "Trading Company Name"
        Range("D1").FormulaR1C1 = "Counterparty"
        Range("E1").FormulaR1C1 = "Counterparty name"
        Range("F1").FormulaR1C1 = "Contract number"
        Range("G1").FormulaR1C1 = "Doc Type"
        Range("H1").FormulaR1C1 = "Deal Type"
        
     
   'Removing duplicate items (double debit or credit)
       With wsMerged
           .Range("A:H").RemoveDuplicates Columns:=Array(6), Header:=xlYes
       End With
  
  
  
'Color formatting header
    Application.StatusBar = "Color formatting header"
        Columns("A:H").Select
        With Selection
            .HorizontalAlignment = xlCenter
        End With
        Range("A1:H1").Select
            Selection.Interior.ColorIndex = 45
        Range("A1:H1").Select
            Selection.Font.Bold = True
        With Selection
            .WrapText = True
            .VerticalAlignment = xlCenter
        End With
        
    Range("A:A").Select
    Selection.Delete Shift:=xlToLeft
        
'color format cells except headers
    
    With wsMerged
    
    Set dataRange = .Range(.Cells(2, 1), .Cells(.UsedRange.Rows.Count, .UsedRange.Columns.Count))
    
    dataRange.Interior.Color = RGB(144, 238, 144)
    
    End With
    


  ' Set up colums width and alignment
    Columns("A:A").ColumnWidth = 18
    Columns("B:B").ColumnWidth = 18
    Columns("C:C").ColumnWidth = 20
    Columns("D:D").ColumnWidth = 19
    Columns("E:E").ColumnWidth = 42
    Columns("F:F").ColumnWidth = 25
    Columns("G:G").ColumnWidth = 25
    Columns("H:H").ColumnWidth = 18
    
    
'close the selected workbook without saving

 ' Workbooks(selectedWorkbook).Close SaveChanges:=False
 
  ' Freeze panels
  Rows("2:2").Select
  ActiveWindow.FreezePanes = True
    
    
     ' Adding filter
  
  Rows("1:1").Select
  Selection.AutoFilter
   
   'close the sourceworkbook
    'sourceWorkbook.Close SaveChanges:=False
   
    End If
End Sub








