Attribute VB_Name = "ProccessData"

Sub ProccessData()
    ' the excel worksheet with the user controls
    Dim wsControls As Worksheet
    Dim numRowsToDeleteFromTop As Variant
    Dim wsSource As Worksheet
    Dim wsDest As Worksheet
    Dim ws As Worksheet
    Dim lastRowSource As Long
    Dim lastRowDest As Long
    Dim lastCol As Long
    Dim tblRange As Range
    Dim tbl As ListObject
    Dim headers As Variant
    Dim destHeaders As Variant
    Dim i As Long
    Dim rowNum As Long
    Dim colNum As Long
    Dim found As Boolean
    Dim cell As Range
    Dim lines() As String
    Dim newContent As String
    Dim cellValue As String
    Dim colShift As Long
    
    On Error GoTo ErrorHandler
    
    ' Headers to use
    headers = Array("Test Name", "Result", "Units", "Reference Range", "Date", "Status")
    destHeaders = Array("Test Name", "Result", "Units", "Reference Range", "Date", "Status")
    
    ' Call the function to find the latest "All Pages" sheet
    Set wsSource = FindLatestAllPagesSheet
    ' find the controls sheet
    Set wsControls = ThisWorkbook.Sheets("Controls")
    
    ' Check if the All Pages sheet was found
    If wsSource Is Nothing Then
        MsgBox "No sheet with 'All Pages' found.", vbExclamation
    Else
        MsgBox "The selected 'All Pages' sheet is: " & wsSource.Name, vbInformation
    End If
    
    ' check if the controls sheet was found
    If wsControls Is Nothing Then
        MsgBox "No sheet with 'Controls' found.", vbExclamation
    Else
        MsgBox "Found the Controls sheet to use for user variable input"
    End If
    
    ' Retrieve the # of rows to delete from the top of the data file from the control sheet
    On Error Resume Next
    numRowsToDeleteFromTop = wsControls.Range("E17").Value
    On Error GoTo 0
    
    ' Rename headers of the source data to the given headers
    For i = 1 To UBound(headers) + 1
        If i <= wsSource.Columns.Count Then
            wsSource.Cells(1, i).Value = headers(i - 1)
        End If
    Next i
    
    'handle clearing the top rows
    'check if the variable is a valid row number
    If IsNumeric(numRowsToDeleteFromTop) And numRowsToDeleteFromTop > 0 Then
        success = ClearFirstRowsContents(wsSource, CInt(numRowsToDeleteFromTop))
    Else
        MsgBox "Invalid input for Step 8: cell E17 please enter a positive number.", vbExclamation
    End If
    
    'check if the clearing was successful
    If success Then
        MsgBox "Rows cleared."
    Else
        MsgBox "No rows cleared."
    End If
    
    ' Find the last row and column with data in the source sheet
    lastRowSource = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row
    lastCol = wsSource.Cells(1, wsSource.Columns.Count).End(xlToLeft).Column
    
    ' Delete empty rows and rows with only invisible characters
    DeleteEmptyRows wsSource
    
    ' Find the last row and column with data in the source sheet
    lastRowSource = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row
    lastCol = wsSource.Cells(1, wsSource.Columns.Count).End(xlToLeft).Column
    
    'clear columns 7 and beyond
    successC = ClearColumnsFromG(wsSource)
    
    'check if the clearing was successful
    If successC Then
        MsgBox "Columns cleared."
    Else
        MsgBox "No Columns cleared."
    End If
    
    ' Delete empty columns and columns with only invisible characters
    For colNum = lastCol To 1 Step -1
        Dim columnIsEmpty As Boolean
        columnIsEmpty = True
        
        For Each cell In wsSource.Columns(colNum).Cells
            If Not isEmpty(cell) And Not ContainsOnlyInvisibleChars(CStr(cell.Value)) Then
                columnIsEmpty = False
                Exit For
            End If
        Next cell
        
        If columnIsEmpty Then
            wsSource.Columns(colNum).Delete
        End If
    Next colNum
    
    ' Delete all hyperlinks in the source sheet
    wsSource.Hyperlinks.Delete
    
    ' Delete rows that contain specific values
    For rowNum = lastRowSource To 1 Step -1
        found = False
        For colNum = 1 To lastCol
            If wsSource.Cells(rowNum, colNum).Value = "MyVirginiaMason - Lab Results" Or _
               wsSource.Cells(rowNum, colNum).Value = "View all for this result" Or _
               InStr(wsSource.Cells(rowNum, colNum).Value, "Differential:") > 0 Or _
               InStr(wsSource.Cells(rowNum, colNum).Value, "https") > 0 Or _
               wsSource.Cells(rowNum, colNum).Value = "Done" Or _
               wsSource.Cells(rowNum, colNum).Value = "General Chemistry" Or _
               wsSource.Cells(rowNum, colNum).Value = "Tumor Markers" Then
                found = True
                Exit For
            End If
        Next colNum
        If found Then
            On Error Resume Next
            wsSource.Rows(rowNum).Delete
            If Err.Number <> 0 Then
                Debug.Print "Error deleting row " & rowNum & ": " & Err.Description
                Err.Clear
            End If
            On Error GoTo ErrorHandler
        End If
    Next rowNum
    
    ' Find the last used row and column after deleting rows
    lastRowSource = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row
    lastCol = wsSource.Cells(1, wsSource.Columns.Count).End(xlToLeft).Column
    
    ' Move data to the left for each row until only empty columns are at the rightmost end
    For rowNum = 1 To lastRowSource
        colNum = 1
        Do While colNum <= lastCol
            If wsSource.Cells(rowNum, colNum).Value = "" Then
                ' Shift data from column colNum+1 to colNum
                For colShift = colNum + 1 To lastCol
                    If wsSource.Cells(rowNum, colShift).Value <> "" Then
                        wsSource.Cells(rowNum, colNum).Value = wsSource.Cells(rowNum, colShift).Value
                        wsSource.Cells(rowNum, colShift).Value = ""
                        Exit For
                    End If
                Next colShift
            End If
            colNum = colNum + 1
        Loop
    Next rowNum
    
    ' Delete empty columns and columns with only invisible characters
    For colNum = lastCol To 1 Step -1
        columnIsEmpty = True
        
        For Each cell In wsSource.Columns(colNum).Cells
            If Not isEmpty(cell) And Not ContainsOnlyInvisibleChars(CStr(cell.Value)) Then
                columnIsEmpty = False
                Exit For
            End If
        Next cell
        
        If columnIsEmpty Then
            wsSource.Columns(colNum).Delete
        End If
    Next colNum
    
    ' Find the last used row and column after deleting rows
    lastRowSource = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row
    lastCol = wsSource.Cells(1, wsSource.Columns.Count).End(xlToLeft).Column
    
    ' Move rows starting with numbers
    MoveRowsStartingWithNumbers wsSource
    
    'Lab results were processed successfully
    MsgBox "Rows Starting With Number Moved."
    
    ' Recalculate lastRowSource after processing
    lastRowSource = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row
    
    ' Move rows starting with Date:
    For rowNum = lastRowSource To 2 Step -1
        If Left(wsSource.Cells(rowNum, 1).Value, 5) = "Date:" Then
            ' Move the first 4 columns to the row above, starting in the third column
            wsSource.Cells(rowNum - 1, 3).Resize(1, 4).Value = wsSource.Cells(rowNum, 1).Resize(1, 4).Value
            ' Clear the original row
            wsSource.Cells(rowNum, 1).Resize(1, 4).ClearContents
            ' If the entire row is empty after clearing, delete it
            If Application.WorksheetFunction.CountA(wsSource.Rows(rowNum)) = 0 Then
                wsSource.Rows(rowNum).Delete
                lastRowSource = lastRowSource - 1
            End If
        End If
    Next rowNum
    
    'Lab results were processed successfully
    MsgBox "Rows Starting With 'Date:' Moved."
    
    ' Recalculate lastRowSource after processing
    lastRowSource = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row
    
    ' Remove "View all for this result" text and empty lines
    For Each cell In wsSource.Range(wsSource.Cells(1, 1), wsSource.Cells(lastRowSource, lastCol))
        If InStr(cell.Value, "View all for this result") > 0 Then
            ' Split the cell content into lines
            lines = Split(cell.Value, vbNewLine)
            newContent = ""
            ' Process each line
            For i = LBound(lines) To UBound(lines)
                ' Remove "View all for this result" from the line
                lines(i) = Replace(lines(i), "View all for this result", "")
                ' If the line is not empty after removal, add it to the new content
                If Trim(lines(i)) <> "" Then
                    If newContent <> "" Then
                        newContent = newContent & vbNewLine
                    End If
                    newContent = newContent & Trim(lines(i))
                End If
            Next i
            ' Update the cell with the new content
            cell.Value = newContent
        End If
    Next cell
    
    ' Delete empty rows and rows with only invisible characters
    DeleteEmptyRows wsSource
    
    ' Recalculate lastRowSource and lastCol after processing
    lastRowSource = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row
    lastCol = wsSource.Cells(1, wsSource.Columns.Count).End(xlToLeft).Column
    
    ' Handle cells with multiple lines
    If ProcessCellsWithMultipleLines(wsSource) Then
        ' Lab results were processed successfully
        MsgBox "Lab results with multiple lines: Processing Complete."
    End If
    
    ' Recalculate lastRowSource and lastCol after processing
    lastRowSource = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row
    lastCol = wsSource.Cells(1, wsSource.Columns.Count).End(xlToLeft).Column
    
    ' Delete empty rows and rows with only invisible characters
    DeleteEmptyRows wsSource
    
    ' Recalculate lastRowSource after deleting rows
    lastRowSource = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row
    
    ' Call the last row manipulating function
    ProcessLabResultRows
    ' processed rows was a success
    MsgBox "Lab results rows: Processing Complete."
    
    ' Recalculate lastRowSource after deleting rows
    lastRowSource = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row
    
    ' Add the names for the Date and Status Columns
    RenameLastTwoColumns
    
   ' Check if a sheet named "Data" already exists
On Error Resume Next
Set wsDest = ThisWorkbook.Sheets("Data")
If Err.Number <> 0 Then
    Debug.Print "Error finding Data sheet: " & Err.Description
    Err.Clear
End If
On Error GoTo ErrorHandler

' If the "Data" sheet does not exist, or if it exists but is empty, create or adjust headers
If wsDest Is Nothing Then
    Set wsDest = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsDest.Name = "Data"
    ' Add headers to the new "Data" sheet
    For i = LBound(destHeaders) To UBound(destHeaders)
        wsDest.Cells(1, i + 1).Value = destHeaders(i)
    Next i
    lastRowDest = 1
ElseIf wsDest.Cells(1, 1).Value = "" Then
    ' Copy headers from the source sheet to the destination sheet if it's empty
    wsSource.Rows(1).Copy Destination:=wsDest.Rows(1)
    lastRowDest = 1
Else
    ' Find the last row in the existing "Data" sheet
    lastRowDest = wsDest.Cells(wsDest.Rows.Count, "A").End(xlUp).Row
End If
    
    ' Copy headers from the source sheet to the destination sheet if it's empty
    If lastRowDest = 0 Then
        wsSource.Rows(1).Copy Destination:=wsDest.Rows(1)
        lastRowDest = 1
    End If
    
    ' Copy data from the source sheet to the destination sheet, starting from the next empty row
    wsSource.Range(wsSource.Cells(2, 1), wsSource.Cells(lastRowSource, lastCol)).Copy
    wsDest.Cells(lastRowDest + 1, 1).PasteSpecial xlPasteAll
    
    ' Set the range for the table in the destination sheet
    Set tblRange = wsDest.Range(wsDest.Cells(1, 1), wsDest.Cells(lastRowDest + lastRowSource - 1, lastCol))
    
    ' Check if a table already exists
    On Error Resume Next
    Set tbl = wsDest.ListObjects("DataTable")
    If Err.Number <> 0 Then
        Debug.Print "Error finding existing DataTable: " & Err.Description
        Err.Clear
    End If
    On Error GoTo ErrorHandler
    
    ' If the table does not exist, create it
    If tbl Is Nothing Then
        Set tbl = wsDest.ListObjects.Add(xlSrcRange, tblRange, , xlYes)
        tbl.Name = "DataTable"
    Else
        ' Resize the existing table to include the new data
        tbl.Resize tblRange
    End If
    
    ' Apply a table style
    tbl.TableStyle = "TableStyleMedium9"
    
    MsgBox "Data appended successfully to the 'Data' sheet."

ErrorHandler:
    Debug.Print "Error: " & Err.Description
    Err.Clear
End Sub
Function FindLatestAllPagesSheet() As Worksheet
    Dim ws As Worksheet
    Dim wsSource As Worksheet
    Dim latestDate As Date
    Dim sheetDate As Date
    Dim datePattern As String
    Dim sheetName As String
    Dim foundSheetCount As Integer
    
    ' Regular expression pattern to extract date (assuming format MM-DD-YYYY)
    datePattern = "(\d{2}-\d{2}-\d{4})"
    
    foundSheetCount = 0 ' Initialize counter
    
    ' Loop through all sheets
    For Each ws In ThisWorkbook.Sheets
        ' Check if the sheet name contains "All Pages" (case-insensitive)
        If InStr(1, ws.Name, "All Pages", vbTextCompare) > 0 Then
            foundSheetCount = foundSheetCount + 1 ' Increment count of found sheets
            ' Extract the date from the sheet name
            sheetName = ws.Name
            On Error Resume Next
            sheetDate = CDate(ExtractDate(sheetName, datePattern))
            On Error GoTo 0
            
            ' If no date is found, consider this sheet if it's the only one found
            If sheetDate = 0 And foundSheetCount = 1 Then
                Set wsSource = ws
            ElseIf sheetDate > latestDate Then
                latestDate = sheetDate
                Set wsSource = ws
            End If
        End If
    Next ws
    
    ' Return the found sheet (wsSource), which may be Nothing if no sheet was found
    Set FindLatestAllPagesSheet = wsSource
End Function

' Function to extract the date from the sheet name based on a pattern
Function ExtractDate(sheetName As String, datePattern As String) As String
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    With regex
        .Pattern = datePattern
        .Global = False
    End With
    
    If regex.test(sheetName) Then
        ExtractDate = regex.Execute(sheetName)(0)
    Else
        ExtractDate = ""
    End If
End Function
'checks if the string as any invisible chars
Function ContainsInvisibleChar(ByVal str As String) As Boolean
    Dim i As Integer
    Dim charCode As Integer
    
    For i = 1 To Len(str)
        charCode = Asc(Mid(str, i, 1))
        ' Check if the character is an invisible character and not a space
        If (charCode >= 0 And charCode <= 32) And charCode <> 32 Then
            ContainsInvisibleChar = True
            Exit Function
        End If
    Next i
    
    ContainsInvisibleChar = False
End Function
'adds the names to the last two columns
Sub RenameLastTwoColumns()
    Dim ws As Worksheet
    Dim lastColumn As Long
    
    ' Set the worksheet you want to modify
    Set ws = ActiveSheet ' Or use a specific sheet, e.g., ThisWorkbook.Sheets("Sheet1")
    
    ' Find the last column with data
    lastColumn = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' Rename the last two columns
    ws.Cells(1, lastColumn - 1).Value = "Date"
    ws.Cells(1, lastColumn).Value = "Status"
End Sub
'used to process the rows with only one line
Sub ProcessLabResultRows()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim resultAndUnits As Variant
    
    ' Assuming we're working with the active sheet
    Set ws = ActiveSheet
    
    ' Find the last row with data
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Loop through each row from bottom to top
    For i = lastRow To 2 Step -1 ' Assuming row 1 is headers
        ' Check if the result is empty or contains only invisible characters
        If Len(Trim(ws.Cells(i, 2).Value)) = 0 Then
    Debug.Print "Error: Empty result for test '" & ws.Cells(i, 1).Value & "' in row " & i
    ws.Rows(i).Delete
        ElseIf ContainsOnlyInvisibleChars(ws.Cells(i, 2).Value) Then
            Debug.Print "Warning: Possibly invisible result for test '" & ws.Cells(i, 1).Value & "' in row " & i & ". Value: '" & ws.Cells(i, 2).Value & "'"
            ' Instead of deleting, you might want to flag these rows for manual review
            ws.Cells(i, 2).Interior.Color = vbYellow
        ' Check if the result is "Done", "1+", see, or nonreactive
        ElseIf ws.Cells(i, 2).Value = "Done" Or ws.Cells(i, 2).Value = "1+" Or LCase(ws.Cells(i, 2).Value) = "see report" Or LCase(ws.Cells(i, 2).Value) = "nonreactive" Then
            Debug.Print "Error: Result is '" & ws.Cells(i, 2).Value & "' for test '" & ws.Cells(i, 1).Value & "' in row " & i
            ws.Rows(i).Delete
        Else
            ' Step 1: Check and separate units in column B if needed
            If InStr(ws.Cells(i, 2).Value, " ") > 0 Or InStr(ws.Cells(i, 2).Value, "%") > 0 Or InStr(ws.Cells(i, 2).NumberFormat, "%") > 0 Then
                resultAndUnits = SeparateUnitsAndResult(ws.Cells(i, 2), i)
                ws.Cells(i, 2).Value = resultAndUnits(0) ' Result
                ws.Cells(i, 3).Value = resultAndUnits(1) ' Units
            ElseIf ws.Cells(i, 3).Value = "" Then
                ' If there's no space or % in the result and no unit in column C, set unit to "NA"
                ws.Cells(i, 3).Value = "NA"
                Debug.Print "Row " & i & ": No unit or % found for result: " & ws.Cells(i, 2).Value
            End If
            
            ' Step 2: Check if column C only has "Date:" or "Date: " and replace with units if needed
            If ws.Cells(i, 3).Value = "Date:" Or ws.Cells(i, 3).Value = "Date: " Then
                ws.Cells(i, 3).Value = resultAndUnits(1) ' Units
            End If
            
            ' Step 3: Check if column D is a date ending with PDT
            If InStr(ws.Cells(i, 4).Value, "PDT") > 0 Then
                ' Step 4: Check if column E only has "Reference Range:" or "Reference Range: "
                If ws.Cells(i, 5).Value = "Reference Range:" Or ws.Cells(i, 5).Value = "Reference Range: " Then
                    ws.Cells(i, 5).Value = ws.Cells(i, 4).Value ' Move date to column E
                    ws.Cells(i, 4).Value = ws.Cells(i, 6).Value ' Move reference range to column D
                    ws.Cells(i, 6).Value = "" ' Clear column F
                ' if there is no reference range
                Else
                    ' Move date to column E
                    ws.Cells(i, 5).Value = ws.Cells(i, 4).Value
                    ' Put "Not Provided" in place of ref range
                    ws.Cells(i, 4).Value = "Not Provided"
                    'Make Status NA
                    ws.Cells(i, 6).Value = "NA"
                End If
            End If
            
            'if there is no date remove reference range and replace with the date from above
            If ws.Cells(i, 5).Value = "Reference Range:" Or ws.Cells(i, 5).Value = "Reference Range: " Then
                ' Move date from row above to column E
                ws.Cells(i, 5).Value = ws.Cells(i + 1, 5).Value
            End If
            
            ' Step 5: Check if the result is within the reference range and set status
            If ws.Cells(i, 4).Value <> "" And ws.Cells(i, 4).Value <> "Not Provided" Then
                ws.Cells(i, 6).Value = DetermineStatus(ws.Cells(i, 2).Value, ws.Cells(i, 4).Value)
            ElseIf Trim(ws.Cells(i, 5).Value) = "" Or ContainsOnlyInvisibleChars(ws.Cells(i, 5).Value) Then
                ' If column E is empty or contains only invisible characters
                ws.Cells(i, 5).Value = ws.Cells(i, 4).Value ' Move date to column E
                ws.Cells(i, 4).Value = "Not Provided"
                ws.Cells(i, 6).Value = "NA"
            End If
        End If
    Next i
End Sub
'Checks cells for multiple lines and processes them by calling the correct functions
Function ProcessCellsWithMultipleLines(ws As Worksheet) As Boolean
    Dim lastRow As Long
    Dim cell As Range
    Dim processedCount As Long
    
    ' Find the last row with data
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    processedCount = 0
    
    ' Loop through each cell in columns A and B
    For Each cell In ws.Range("A1:B" & lastRow)
        ' Check if the cell contains multiple lines
        If HasMultipleLines(cell) Then
            ' Try to separate the lab result
            If SeparateLabResult(cell) Then
                processedCount = processedCount + 1
            Else
                ' If separation fails, you might want to log this or handle it somehow
                Debug.Print "Failed to separate lab result in row " & cell.Row & ", column " & cell.Column
            End If
        End If
    Next cell
    
    ' Return True if any cells were processed, False otherwise
    ProcessCellsWithMultipleLines = (processedCount > 0)
End Function
'used to split up the cells with multiple lines
Function SeparateLabResult(cell As Range) As Boolean
    Dim lines() As String
    Dim i As Long
    Dim testName As String, result As String, units As String, status As String
    Dim testDate As String, refRange As String
    Dim cleanLine As String
    
    ' Split the cell content into lines using both possible line break characters
    lines = Split(Replace(Replace(cell.Value, Chr(10), "|"), Chr(13), "|"), "|")
    
    ' Initialize variables
    testName = ""
    result = ""
    units = ""
    status = ""
    testDate = ""
    refRange = ""
    
    ' Process lines
    For i = 0 To UBound(lines)
        cleanLine = Trim(lines(i))
        
        ' Test Name (first non-empty line)
        If testName = "" And cleanLine <> "" Then
            testName = cleanLine
        ' Result, Units, and Status
        ElseIf (IsNumeric(Left(cleanLine, 1)) Or Left(cleanLine, 1) = "<" Or Left(cleanLine, 1) = ">" Or IsNumeric(Mid(cleanLine, 2, 1))) And result = "" Then
            Dim resultArray As Variant
            ' Create a temporary cell to hold the cleanLine value
            Dim tempCell As Range
            Set tempCell = cell.Worksheet.Cells(cell.Row, cell.Column)
            tempCell.Value = cleanLine
            resultArray = SeparateUnitsAndResult(tempCell, cell.Row)
            If IsArray(resultArray) And UBound(resultArray) >= 2 Then
                result = CStr(resultArray(0))
                units = CStr(resultArray(1))
                If Len(CStr(resultArray(2))) > 0 Then
                    status = CStr(resultArray(2))
                End If
            End If
        ' Date and Reference Range
        ElseIf InStr(cleanLine, "Date:") > 0 And InStr(cleanLine, "PDT") > 0 Then
            Dim dateEndPos As Long
            dateEndPos = InStr(cleanLine, "PDT") + 3
            testDate = Trim(Mid(cleanLine, InStr(cleanLine, ":") + 1, dateEndPos - InStr(cleanLine, ":") - 1))
            If InStr(cleanLine, "Reference Range:") > 0 Then
                Dim refRangePart As String
                refRangePart = Trim(Mid(cleanLine, InStr(cleanLine, "Reference Range:") + 16))
                ' Check if the reference range part is empty or contains only invisible characters
                If refRangePart = "" Or ContainsOnlyInvisibleChars(refRangePart) Then
                    ' Use the next line as the reference range if available
                    If i < UBound(lines) Then
                        refRange = Trim(lines(i + 1))
                    End If
                Else
                    refRange = refRangePart
                End If
            End If
        End If
    Next i
    
    ' Determine status if not already set
    If status = "" Then
        If refRange <> "" Then
            Dim lowRange As Double, highRange As Double
            Dim resultValue As Double
            ' Extract low and high range values, ignoring units and special characters
            Dim rangeParts() As String
            rangeParts = Split(refRange, "-")
            If UBound(rangeParts) = 1 Then
                lowRange = ExtractNumericValue(rangeParts(0))
                highRange = ExtractNumericValue(rangeParts(1))
                ' Convert result to a number, handling < and >
                If Left(result, 1) = "<" Or Left(result, 1) = ">" Then
                    resultValue = ExtractNumericValue(Mid(result, 2))
                    If resultValue >= lowRange And resultValue <= highRange Then
                        status = "Normal"
                    ElseIf resultValue > highRange Then
                        status = "High"
                    Else
                        status = "Low"
                    End If
                Else
                    resultValue = ExtractNumericValue(result)
                    ' Determine status based on range for regular numbers
                    If resultValue < lowRange Then
                        status = "Low"
                    ElseIf resultValue > highRange Then
                        status = "High"
                    Else
                        status = "Normal"
                    End If
                End If
                If resultValue = 0 Then
                    status = "NA"
                    Debug.Print "Row " & cell.Row & ": NA - Unable to convert result to number: " & result
                End If
            Else
                status = "NA"
                Debug.Print "Row " & cell.Row & ": NA - Invalid reference range format: " & refRange
            End If
        Else
            status = "NA"
            Debug.Print "Row " & cell.Row & ": NA - No reference range provided"
        End If
    End If
    
    ' Write to cells
    cell.Value = testName
    cell.Offset(0, 1).Value = result
    cell.Offset(0, 2).Value = units
    cell.Offset(0, 3).Value = refRange
    cell.Offset(0, 4).Value = testDate
    cell.Offset(0, 5).Value = status
    
    SeparateLabResult = True
End Function
'grab only the numbers in the string
Function ExtractNumericValue(str As String) As Double
    Dim i As Long
    Dim numStr As String
    For i = 1 To Len(str)
        If IsNumeric(Mid(str, i, 1)) Or Mid(str, i, 1) = "." Then
            numStr = numStr & Mid(str, i, 1)
        End If
    Next i
    If numStr <> "" Then
        ExtractNumericValue = CDbl(numStr)
    Else
        ExtractNumericValue = 0
    End If
End Function
'Separarte the units from the result
Function SeparateUnitsAndResult(cell As Range, rowNumber As Long) As Variant
    Dim inputString As String
    Dim result As String
    Dim units As String
    Dim status As String
    Dim parts() As String
    Dim i As Long
    
    ' Get the value and number format of the cell
    inputString = cell.Value
    Dim cellFormat As String
    cellFormat = cell.NumberFormat
    
    ' Remove any leading or trailing spaces
    inputString = Trim(inputString)
    
    ' Check if the cell is formatted as a percentage
    If InStr(cellFormat, "%") > 0 Then
        Debug.Print "Row " & rowNumber & ": Percentage found (Excel format): " & inputString
        result = inputString
        units = "%"
    ' Check for % in the entire input string
    ElseIf InStr(inputString, "%") > 0 Then
        Debug.Print "Row " & rowNumber & ": Percentage found in result: " & inputString
        Dim percentIndex As Long
        percentIndex = InStr(inputString, "%")
        result = Left(inputString, percentIndex)  ' Keep the % with the result
        units = "%"
    Else
        ' Split the input string by spaces
        parts = Split(inputString, " ")
        
        ' Initialize result and units
        result = ""
        units = ""
        
       ' Handle < or > with the number
If UBound(parts) >= 0 Then
    If Left(parts(0), 1) = "<" Or Left(parts(0), 1) = ">" Then
        ' Remove the < or > after joining with the number
        result = Replace(parts(0), "<", "")  ' Remove "<"
        result = Replace(result, ">", "")    ' Remove ">"
        
        If UBound(parts) > 0 Then
            ' Check if the next part is numeric (in case of space after < or >)
            If IsNumeric(parts(1)) Then
                result = result & parts(1)
                If UBound(parts) > 2 Then
                    units = Join(Array(parts(2), parts(3)), " ")
                ElseIf UBound(parts) > 1 Then
                    units = parts(2)
                End If
            Else
                If UBound(parts) > 1 Then
                    units = Join(Array(parts(1), parts(2)), " ")
                ElseIf UBound(parts) > 0 Then
                    units = parts(1)
                End If
            End If
        End If
    Else
        result = parts(0)
        If UBound(parts) > 1 Then
            units = Join(Array(parts(1), parts(2)), " ")
        ElseIf UBound(parts) > 0 Then
            units = parts(1)
        End If
    End If
End If
        
        ' Check for status in parentheses
        If InStr(inputString, "(") > 0 Then
            status = Mid(inputString, InStr(inputString, "("))
            status = Left(status, Len(status) - 1) ' Remove closing parenthesis
        End If
    End If
    
    ' If no unit or % is found, set unit to "NA" and print error message
    If units = "" Or units = "report" Or units = "see" Then
        units = "NA"
        Debug.Print "Row " & rowNumber & ": No unit or % found for result: " & inputString
    End If
    
    ' Trim any extra spaces from result, units, and status
    result = Trim(result)
    units = Trim(units)
    status = Trim(status)
    
    ' Return an array with result, units, and status
    SeparateUnitsAndResult = Array(result, units, status)
End Function
'Calculate the status of the test by checking the result against the reference range
Function DetermineStatus(result As String, refRange As String) As String
    Dim resultValue As Double
    Dim lowRange As Double, highRange As Double
    Dim rangeParts() As String
    
    ' Extract numeric value from result
    resultValue = ExtractNumericValue(result)
    
    ' Extract low and high range values
    rangeParts = Split(refRange, "-")
    If UBound(rangeParts) = 1 Then
        lowRange = ExtractNumericValue(rangeParts(0))
        highRange = ExtractNumericValue(rangeParts(1))
        
        ' Determine status
        If Left(result, 1) = "<" Then
            If resultValue <= highRange Then
                DetermineStatus = "Normal"
            Else
                DetermineStatus = "High"
            End If
        ElseIf Left(result, 1) = ">" Then
            If resultValue >= lowRange Then
                DetermineStatus = "Normal"
            Else
                DetermineStatus = "Low"
            End If
        Else
            If resultValue < lowRange Then
                DetermineStatus = "Low"
            ElseIf resultValue > highRange Then
                DetermineStatus = "High"
            Else
                DetermineStatus = "Normal"
            End If
        End If
    Else
        DetermineStatus = "NA"
    End If
End Function
'Check if there are only invisible chars in the string
Function ContainsOnlyInvisibleChars(ByVal str As String) As Boolean
    Dim i As Long
    Dim charCode As Long
    For i = 1 To Len(str)
        charCode = AscW(Mid(str, i, 1))
        ' Check for visible characters (including spaces)
        If (charCode > 32 And charCode <> 160) Or charCode = 32 Then
            ContainsOnlyInvisibleChars = False
            Exit Function
        End If
    Next i
    ContainsOnlyInvisibleChars = True
End Function
'Check if a cell has multiple lines by looking at the height of the cell - checking for endline chars doesn't work
Function HasMultipleLines(cell As Range) As Boolean
    Dim textHeight As Double
    Dim cellHeight As Double
    
    ' Get the height of the cell's text if it were on one line
    textHeight = cell.Font.Size * 1.3 ' 1.3 is an approximation for line height
    
    ' Get the actual height of the cell
    cellHeight = cell.RowHeight
    
    ' If the cell height is significantly larger than the text height, it likely contains multiple lines
    HasMultipleLines = (cellHeight > (textHeight * 1.5))
End Function
'Delete rows that are empty or only have invisible chars
Sub DeleteEmptyRows(ws As Worksheet)
    Dim lastRow As Long
    Dim rowNum As Long

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    For rowNum = lastRow To 1 Step -1
        If Application.WorksheetFunction.CountA(ws.Rows(rowNum)) = 0 Or _
           ContainsOnlyInvisibleChars(Join(Application.Transpose(Application.Transpose(ws.Rows(rowNum).Value)))) Then
            On Error Resume Next
            ws.Rows(rowNum).Delete
            If Err.Number <> 0 Then
                Debug.Print "Error deleting row " & rowNum & ": " & Err.Description
                Err.Clear
            End If
            On Error GoTo 0
        End If
    Next rowNum
End Sub
'Delete columns that are empty or only have invisible chars
Sub DeleteEmptyColumns(ws As Worksheet)
    Dim lastCol As Long
    Dim colNum As Long

    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    For colNum = lastCol To 1 Step -1
        If Application.WorksheetFunction.CountA(ws.Columns(colNum)) = 0 Or _
           ContainsOnlyInvisibleChars(Join(Application.Transpose(Application.Transpose(ws.Columns(colNum).Value)), "")) Then
            On Error Resume Next
            ws.Columns(colNum).Delete
            If Err.Number <> 0 Then
                Debug.Print "Error deleting column " & colNum & ": " & Err.Description
                Err.Clear
            End If
            On Error GoTo 0
        End If
    Next colNum
End Sub
Function ClearFirstRowsContents(wsSource As Worksheet, numRows As Integer) As Boolean
    On Error GoTo ErrorHandler

    ' Unprotect the worksheet if it's protected
    If wsSource.ProtectContents Then
        wsSource.Unprotect ' Add password if needed
    End If

    ' Check if the specified number of rows is valid
    If numRows > wsSource.Rows.Count Then
        MsgBox "The specified number of rows exceeds the total number of rows in the sheet.", vbExclamation
        ClearFirstRowsContents = False
        Exit Function
    End If

    ' Clear the contents of the first numRows rows
    wsSource.Rows("1:" & numRows).ClearContents
    
    ' Inform the user
    MsgBox "Cleared the contents of the first " & numRows & " rows in sheet: " & wsSource.Name, vbInformation
    ClearFirstRowsContents = True ' Return True if clearing was successful
    Exit Function

ErrorHandler:
    MsgBox "An error occurred while trying to clear the contents: " & Err.Description, vbCritical
    ClearFirstRowsContents = False ' Return False if an error occurred
End Function
Function ClearColumnsFromG(wsSource As Worksheet) As Boolean
    On Error GoTo ErrorHandler

    ' Unprotect the worksheet if it's protected
    If wsSource.ProtectContents Then
        wsSource.Unprotect ' Add password if needed
    End If

    ' Clear contents of columns from G (7) and beyond
    Dim lastCol As Long
    lastCol = wsSource.Columns.Count

    ' Clear columns from G (7) to the last column
    wsSource.Range(wsSource.Cells(1, 7), wsSource.Cells(1, lastCol)).EntireColumn.ClearContents
    
    ' Inform the user
    MsgBox "Cleared contents of columns from column G and beyond in sheet: " & wsSource.Name, vbInformation
    ClearColumnsFromG = True ' Return True if clearing was successful
    Exit Function

ErrorHandler:
    MsgBox "An error occurred while trying to clear the columns: " & Err.Description, vbCritical
    ClearColumnsFromG = False ' Return False if an error occurred
End Function
Sub MoveRowsStartingWithNumbers(ws As Worksheet)
    Dim lastRow As Long
    Dim rowNum As Long
    Dim colNum As Long
    Dim cellValue As String
    Dim hasVisibleChars As Boolean
    Dim firstChar As String

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    For rowNum = lastRow To 2 Step -1
        firstChar = Left(Trim(ws.Cells(rowNum, 1).Value), 1)
        If IsNumeric(firstChar) Or _
           firstChar = ">" Or _
           firstChar = "<" Or _
           LCase(Trim(ws.Cells(rowNum, 1).Value)) = "nonreactive" Or _
           LCase(Trim(ws.Cells(rowNum, 1).Value)) = "see report" Then
            
            ws.Cells(rowNum - 1, 2).Value = ws.Cells(rowNum, 1).Value
            ws.Cells(rowNum, 1).ClearContents
            
            hasVisibleChars = False
            For colNum = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
                cellValue = Trim(ws.Cells(rowNum, colNum).Value)
                If cellValue <> "" And Not ContainsOnlyInvisibleChars(cellValue) Then
                    hasVisibleChars = True
                    Exit For
                End If
            Next colNum
            
            If Not hasVisibleChars Then
                ws.Rows(rowNum).Delete
                lastRow = lastRow - 1
            End If
        End If
    Next rowNum
End Sub

