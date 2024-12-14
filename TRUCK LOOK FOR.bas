Attribute VB_Name = "mod_lookfor"

Public Function lookfor(wrksht As Worksheet, strInput As String) As Variant

Dim col As Integer
col = 1
While wrksht.Cells(1, col) <> ""
    If wrksht.Cells(1, col) = strInput Then
        lookfor = col
        GoTo skiptherest
    End If
col = col + 1
Wend

skiptherest:
End Function

Public Function test()

Dim wrkbk As Workbook
Dim wrksht As Worksheet

Set wrkbk = Excel.Application.Workbooks("Truck Project.xlsm")
Set wrksht = wrkbk.Sheets("INPUT")

Debug.Print lookfor(wrksht, "LineCompanyCode")

End Function
