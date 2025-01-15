Attribute VB_Name = "DeleteAllUnitSheets"
Sub DeleteAllUnitSheets()
    Dim ws As Worksheet
    Dim wsToDelete As Worksheet
    Dim i As Long
    Dim sheetCount As Long
    Dim deletedCount As Long
    
    ' Disable screen updating to improve performance
    Application.ScreenUpdating = False
    
    ' Count total sheets
    sheetCount = ThisWorkbook.Sheets.Count
    deletedCount = 0
    
    ' Loop through all sheets in reverse order
    For i = sheetCount To 1 Step -1
        Set ws = ThisWorkbook.Sheets(i)
        
        ' Check if the sheet name is not one of the protected names
        If ws.Name <> "Data" And _
           ws.Name <> "All Graphs" And _
           ws.Name <> "All pages" And _
           Left(ws.Name, 5) <> "Sheet" Then  ' Avoid deleting default Excel sheet names
            
            ' Set the sheet to delete
            Set wsToDelete = ws
            
            ' Delete the sheet
            Application.DisplayAlerts = False  ' Suppress the delete confirmation
            wsToDelete.Delete
            Application.DisplayAlerts = True
            
            deletedCount = deletedCount + 1
        End If
    Next i
    
    ' Re-enable screen updating
    Application.ScreenUpdating = True
    
    ' Show message with results
    MsgBox deletedCount & " unit sheets have been deleted.", vbInformation
End Sub


