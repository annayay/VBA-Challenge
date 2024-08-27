Attribute VB_Name = "Module1"
Sub Reset_Button()
    Dim ws As Worksheet
    Dim c As Long
    Dim x As Long
    Dim lastRow As Long

    For Each ws In ThisWorkbook.Worksheets
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        For x = 1 To lastRow
            For c = 9 To 17 ' Assuming you want to clear values in columns A to D
                ws.Cells(x, c).Value = ""
                ws.Cells(x, c).Interior.Color = xlNone
            Next c
        Next x
    Next ws
End Sub
