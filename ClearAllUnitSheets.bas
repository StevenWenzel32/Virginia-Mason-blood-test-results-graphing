Attribute VB_Name = "ClearAllUnitSheets"
Sub ClearAllUnitSheets()
    Dim ws As Worksheet
    Dim dataSheet As Worksheet
    Dim tbl As ListObject
    
    ' Loop through all sheets
    For Each ws In ThisWorkbook.Worksheets
        ' Check if the sheet name could be a unit (not "Data", "All Graphs", etc.)
        If ws.Name <> "Data" And ws.Name <> "All Graphs" And ws.Name <> "All pages" Then
            ' Clear the contents of the sheet
            ws.Cells.Clear
            ' Delete any tables in the sheet
            For Each tbl In ws.ListObjects
                tbl.Delete
            Next tbl
        End If
    Next ws
    
    MsgBox "All unit sheets have been cleared and their tables deleted."
End Sub

