Attribute VB_Name = "ClearDataSheet"
Sub ClearDataSheet()
    ' Attempt to find and clear the "Data" sheet
    On Error Resume Next
    Set dataSheet = ThisWorkbook.Sheets("Data")
    On Error GoTo 0
    
    If Not dataSheet Is Nothing Then
        ' Clear the contents of the "Data" sheet
        dataSheet.Cells.Clear
        ' Delete any tables in the "Data" sheet
        For Each tbl In dataSheet.ListObjects
            tbl.Delete
        Next tbl
        MsgBox "Data sheet has been cleared."
    Else
        MsgBox "Data sheet not found."
    End If
End Sub

