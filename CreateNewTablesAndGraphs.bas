Attribute VB_Name = "CreateNewTablesAndGraphs"
Sub CreateNewTablesAndGraphs()
    Dim wsSource As Worksheet
    Dim wsDest As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim unit As Variant
    Dim lastRow As Long
    Dim headerRow As Range
    Dim dict As Object
    Dim ws As Worksheet
    Dim sheetExists As Boolean
    Dim allGraphsSheet As Worksheet
    Dim tableName As String
    Dim cht As ChartObject
    Dim successCount As Integer
    Dim totalSheets As Integer
    Dim cleanUnit As String
    
    ' Set the source worksheet
    Set wsSource = ThisWorkbook.Sheets("Data")
    
    ' Find the last row with data
    lastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row
    
    ' Create or get the "All Graphs" sheet and clear existing charts
    On Error Resume Next
    Set allGraphsSheet = ThisWorkbook.Sheets("All Graphs")
    On Error GoTo 0
    If allGraphsSheet Is Nothing Then
        Set allGraphsSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        allGraphsSheet.Name = "All Graphs"
    Else
        ' Clear existing charts
        For Each cht In allGraphsSheet.ChartObjects
            cht.Delete
        Next cht
    End If
    
    ' Set the header row range
    Set headerRow = wsSource.Rows(1)
    
    ' Create a dictionary to store unique units
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Find the column index for the "Units" column
    Dim unitCol As Long
    unitCol = WorksheetFunction.Match("Units", wsSource.Rows(1), 0)
    
    ' Loop through the data to collect unique units
    For Each cell In wsSource.Range(wsSource.Cells(2, unitCol), wsSource.Cells(lastRow, unitCol))
        cleanUnit = CleanUnitName(cell.Value)
        If Not dict.exists(cleanUnit) Then
            dict.Add cleanUnit, Nothing
        End If
    Next cell
    
    successCount = 0
    totalSheets = dict.Count
    
    ' Loop through each unique unit to create or update sheets
    For Each unit In dict.Keys
        sheetExists = False
        
        ' Replace "/" with "-" in the sheet name
        Dim sheetName As String
        sheetName = Replace(unit, "/", "-")
        
        ' Check if the sheet already exists
        For Each ws In ThisWorkbook.Sheets
            If ws.Name = sheetName Then
                sheetExists = True
                Set wsDest = ws
                Exit For
            End If
        Next ws
        
        ' If the sheet does not exist, add a new sheet
        If Not sheetExists Then
            Set wsDest = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
            On Error Resume Next
            wsDest.Name = sheetName
            If Err.Number <> 0 Then
                MsgBox "Error creating sheet for unit '" & unit & "': " & Err.Description
                Err.Clear
                GoTo NextUnit
            End If
            On Error GoTo 0
            
            ' Copy header row to new sheet
            headerRow.Copy Destination:=wsDest.Rows(1)
        Else
            ' Clear existing table if it exists
            On Error Resume Next
            wsDest.ListObjects(1).Delete
            On Error GoTo 0
            
            ' Copy header row to existing sheet
            headerRow.Copy Destination:=wsDest.Rows(1)
        End If
        
        ' Copy data rows to new sheet
        For Each cell In wsSource.Range(wsSource.Cells(2, unitCol), wsSource.Cells(lastRow, unitCol))
            If CleanUnitName(cell.Value) = unit Then
                cell.EntireRow.Copy Destination:=wsDest.Cells(wsDest.Rows.Count, "A").End(xlUp).Offset(1)
            End If
        Next cell
        
        ' Create table
        tableName = "Table_" & Replace(sheetName, " ", "_")
        On Error Resume Next
        Set rng = wsDest.Range("A1").CurrentRegion
        If wsDest.ListObjects.Count = 0 Then
            wsDest.ListObjects.Add(xlSrcRange, rng, , xlYes).Name = tableName
        Else
            wsDest.ListObjects(1).Resize rng
        End If
        If Err.Number <> 0 Then
            MsgBox "Error creating table in sheet '" & sheetName & "': " & Err.Description
            Err.Clear
            GoTo NextUnit
        End If
        On Error GoTo 0
        
        ' Create or update chart
        Call CreateOrUpdateChart(wsDest, allGraphsSheet)
        
        successCount = successCount + 1
        
NextUnit:
    Next unit
    
    If successCount = totalSheets Then
        MsgBox "Tables and Graphs Created Successfully"
    Else
        MsgBox successCount & " out of " & totalSheets & " tables and graphs created successfully"
    End If
End Sub
Sub CreateOrUpdateChart(ws As Worksheet, allGraphsSheet As Worksheet)
    Dim cht As ChartObject
    Dim rng As Range
    Dim tbl As ListObject
    Dim dateCol As Long, testNameCol As Long, resultCol As Long
    Dim uniqueTests As Collection
    Dim test As Variant
    Dim xValues As Range, yValues As Range
    Dim cell As Range
    Dim checkResult As Range
    Dim rowsToDelete As New Collection
    
    ' Find the table in the worksheet
    If ws.ListObjects.Count = 0 Then
        MsgBox "No table found in sheet '" & ws.Name & "'"
        Exit Sub
    End If
    Set tbl = ws.ListObjects(1)
    
    ' Find the date, test name, and result columns
    On Error Resume Next
    dateCol = WorksheetFunction.Match("Date", tbl.HeaderRowRange, 0)
    testNameCol = WorksheetFunction.Match("Test Name", tbl.HeaderRowRange, 0)
    resultCol = WorksheetFunction.Match("Result", tbl.HeaderRowRange, 0)
    On Error GoTo 0
    
    ' Check if all required columns were found
    If dateCol = 0 Or testNameCol = 0 Or resultCol = 0 Then
        MsgBox "One or more required columns (Date, Test Name, Result) not found in the table in sheet '" & ws.Name & "'. Please check your column headers.", vbExclamation
        Exit Sub
    End If
    
   ' Convert date strings to date values
    For Each cell In tbl.ListColumns(dateCol).DataBodyRange
        cell.Value = ParseDate(cell.Value)
    Next cell
    
    ' Sort the table by the Date column to ensure chronological order
    tbl.Sort.SortFields.Clear
    tbl.Sort.SortFields.Add Key:=tbl.ListColumns(dateCol).Range, SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With tbl.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ' Get unique test names
    Set uniqueTests = New Collection
    On Error Resume Next
    For Each rng In tbl.ListColumns(testNameCol).DataBodyRange
        uniqueTests.Add rng.Value, CStr(rng.Value)
    Next rng
    On Error GoTo 0
    
    ' Clear existing charts in the worksheet
    For Each cht In ws.ChartObjects
        cht.Delete
    Next cht
    
    ' Create a new chart
    Set cht = ws.ChartObjects.Add(Left:=Application.InchesToPoints(5), Width:=Application.InchesToPoints(8), Top:=10, Height:=Application.InchesToPoints(4))
    
    ' Add series to the chart
    For Each test In uniqueTests
        ' Filter the table for the current test
        tbl.Range.AutoFilter Field:=testNameCol, Criteria1:=test
        
        ' Get the visible cells for date and result
        Set xValues = tbl.ListColumns(dateCol).DataBodyRange.SpecialCells(xlCellTypeVisible)
        Set yValues = tbl.ListColumns(resultCol).DataBodyRange.SpecialCells(xlCellTypeVisible)
        
        ' Add the series to the chart
        With cht.Chart.SeriesCollection.NewSeries
            .Name = test
            .xValues = xValues
            .Values = yValues
            .MarkerStyle = xlMarkerStyleCircle
        End With
        
        ' Clear the filter
        tbl.AutoFilter.ShowAllData
    Next test
    
    ' Set chart properties
    With cht.Chart
        .ChartType = xlLineMarkers
        .HasTitle = True
        .ChartTitle.Text = ws.Name
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Text = "Date"
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Text = "Result"
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
        
        ' Adjust the number of marks on the x-axis
        .Axes(xlCategory).TickLabelSpacing = WorksheetFunction.Max(1, WorksheetFunction.RoundUp(tbl.ListColumns(dateCol).DataBodyRange.Rows.Count / 10, 0))
    End With
    
    ' Copy chart to "All Graphs" sheet
    cht.Copy
    With allGraphsSheet
        .Paste
        ' Position and resize pasted chart
        PositionChartOnAllGraphs .ChartObjects(.ChartObjects.Count), allGraphsSheet, .ChartObjects.Count = 1
    End With
End Sub

Function ParseDate(dateString As String) As Date
    Dim dateValue As Date
    Dim cleanedString As String
    
    ' Remove the time zone information
    cleanedString = Split(dateString, " ")(0) & " " & Split(dateString, " ")(1) & " " & Split(dateString, " ")(2)
    
    ' Convert to date
    dateValue = CDate(cleanedString)
    
    ParseDate = dateValue
End Function
Sub PositionChart(cht As ChartObject, tbl As ListObject, ws As Worksheet)
    Dim tblTop As Double
    Dim tblLeft As Double
    Dim tblWidth As Double
    Dim tblHeight As Double
    Dim chartTop As Double
    Dim chartLeft As Double
    Dim chartWidth As Double
    Dim chartHeight As Double
    Dim pasteTop As Double
    Dim pasteLeft As Double
    
    ' Get table dimensions and position
    With tbl.Range
        tblTop = .Cells(1, 1).Top
        tblLeft = .Cells(1, 1).Left
        tblHeight = .Cells(.Rows.Count, .Columns.Count).Top - .Cells(1, 1).Top + .Cells(1, 1).Height
        tblWidth = .Cells(.Rows.Count, .Columns.Count).Left - .Cells(1, 1).Left + .Cells(1, 1).Width
    End With
    
    ' Calculate chart dimensions and position
    chartWidth = Application.InchesToPoints(8) ' Set width to 8 inches
    chartHeight = Application.InchesToPoints(5) ' Set height to 4 inches
    
    ' Calculate paste position
    pasteTop = tblTop
    pasteLeft = tblLeft + tblWidth + 20
    
    ' Position and resize chart
    With cht
        .Top = pasteTop
        .Left = pasteLeft
        .Width = chartWidth
        .Height = chartHeight
    End With
End Sub

Function CleanUnitName(unitName As String) As String
    Dim cleanName As String
    cleanName = Trim(unitName)
    cleanName = Replace(cleanName, "(Low)", "")
    cleanName = Replace(cleanName, "(High)", "")
    cleanName = RemoveInvisibleChars(cleanName)
    CleanUnitName = Trim(cleanName)
End Function

Function RemoveInvisibleChars(ByVal str As String) As String
    Dim i As Long
    Dim result As String
    For i = 1 To Len(str)
        If AscW(Mid(str, i, 1)) > 32 Then
            result = result & Mid(str, i, 1)
        End If
    Next i
    RemoveInvisibleChars = Trim(result)
End Function
Sub PositionChartOnAllGraphs(cht As ChartObject, allGraphsSheet As Worksheet, isFirstChart As Boolean)
    Dim chartTop As Double
    Dim chartLeft As Double
    Dim chartWidth As Double
    Dim chartHeight As Double
    Dim pasteTop As Double
    Dim pasteLeft As Double
    Dim lastChart As ChartObject
    Dim offsetTop As Double
    Dim offsetLeft As Double
    Dim maxTop As Double
    
    ' Set standard chart dimensions
    chartWidth = Application.InchesToPoints(8) ' Set width to 8 inches
    chartHeight = Application.InchesToPoints(5) ' Set height to 5 inches
    
    ' Calculate paste position
    If allGraphsSheet.ChartObjects.Count > 0 Then
        ' Find the maximum top position of existing charts
        maxTop = 0
        For Each lastChart In allGraphsSheet.ChartObjects
            If lastChart.Top + lastChart.Height > maxTop Then
                maxTop = lastChart.Top + lastChart.Height
            End If
        Next lastChart
        
        ' Determine the next position
        If isFirstChart Then
            pasteTop = 0
            pasteLeft = 0
        Else
            pasteTop = maxTop + 20
            pasteLeft = 0
        End If
    Else
        ' Initial position for the first chart
        pasteTop = 0
        pasteLeft = 0
    End If
    
    ' Position and resize chart
    With cht
        .Top = pasteTop
        .Left = pasteLeft
        .Width = chartWidth
        .Height = chartHeight
    End With
End Sub

