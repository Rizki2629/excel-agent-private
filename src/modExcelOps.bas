Attribute VB_Name = "modExcelOps"
'==============================================================================
' Excel Agent Add-in - Excel Operations Module
' Handles Excel data operations, formatting, charts, and export
'==============================================================================

Option Explicit

'==============================================================================
' ReadRange - Read data from a range
'==============================================================================
Public Function ReadRange(sheetName As String, rangeAddress As String) As Variant
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Sheets(sheetName)

    Dim rng As Range
    Set rng = ws.Range(rangeAddress)

    If rng.Cells.Count = 1 Then
        ReadRange = rng.Value
    Else
        ReadRange = rng.Value
    End If

    Exit Function

ErrorHandler:
    ReadRange = CVErr(xlErrRef)
End Function

'==============================================================================
' WriteRange - Write data to a range
'==============================================================================
Public Function WriteRange(sheetName As String, rangeAddress As String, data As Variant) As Boolean
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Sheets(sheetName)

    Dim rng As Range
    Set rng = ws.Range(rangeAddress)

    If IsArray(data) Then
        rng.Resize(UBound(data, 1) - LBound(data, 1) + 1, _
                   UBound(data, 2) - LBound(data, 2) + 1).Value = data
    Else
        rng.Value = data
    End If

    WriteRange = True

    Exit Function

ErrorHandler:
    WriteRange = False
End Function

'==============================================================================
' ClearRange - Clear a range
'==============================================================================
Public Function ClearRange(sheetName As String, rangeAddress As String) As Boolean
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Sheets(sheetName)

    Dim rng As Range
    Set rng = ws.Range(rangeAddress)

    rng.Clear

    ClearRange = True

    Exit Function

ErrorHandler:
    ClearRange = False
End Function

'==============================================================================
' FormatRange - Format a range
'==============================================================================
Public Function FormatRange(sheetName As String, rangeAddress As String, _
                           Optional fontName As String = "", _
                           Optional fontSize As Long = 0, _
                           Optional isBold As Boolean = False, _
                           Optional fontColor As Long = 0, _
                           Optional bgColor As Long = 0, _
                           Optional numberFormat As String = "") As Boolean
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Sheets(sheetName)

    Dim rng As Range
    Set rng = ws.Range(rangeAddress)

    With rng.Font
        If fontName <> "" Then .Name = fontName
        If fontSize > 0 Then .Size = fontSize
        .Bold = isBold
        If fontColor <> 0 Then .Color = fontColor
    End With

    If bgColor <> 0 Then
        rng.Interior.Color = bgColor
    End If

    If numberFormat <> "" Then
        rng.NumberFormat = numberFormat
    End If

    FormatRange = True

    Exit Function

ErrorHandler:
    FormatRange = False
End Function

'==============================================================================
' SetCellValue - Set value of a single cell
'==============================================================================
Public Function SetCellValue(sheetName As String, cellAddress As String, value As Variant) As Boolean
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Sheets(sheetName)

    ws.Range(cellAddress).Value = value

    SetCellValue = True

    Exit Function

ErrorHandler:
    SetCellValue = False
End Function

'==============================================================================
' GetCellValue - Get value of a single cell
'==============================================================================
Public Function GetCellValue(sheetName As String, cellAddress As String) As Variant
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Sheets(sheetName)

    GetCellValue = ws.Range(cellAddress).Value

    Exit Function

ErrorHandler:
    GetCellValue = CVErr(xlErrRef)
End Function

'==============================================================================
' CreateChart - Create a chart
'==============================================================================
Public Function CreateChart(sheetName As String, dataRange As String, _
                           chartType As Long, Optional chartTitle As String = "") As Boolean
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Sheets(sheetName)

    Dim dataRng As Range
    Set dataRng = ws.Range(dataRange)

    Dim cht As ChartObject
    Set cht = ws.ChartObjects.Add( _
        Left:=100, Top:=100, _
        Width:=400, Height:=300)

    With cht.Chart
        .SetSourceData Source:=dataRng
        .ChartType = chartType

        If chartTitle <> "" Then
            .HasTitle = True
            .ChartTitle.Text = chartTitle
        End If

        .HasLegend = True
    End With

    CreateChart = True

    Exit Function

ErrorHandler:
    CreateChart = False
End Function

'==============================================================================
' DeleteChart - Delete a chart
'==============================================================================
Public Function DeleteChart(sheetName As String, Optional chartIndex As Long = 1) As Boolean
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Sheets(sheetName)

    If ws.ChartObjects.Count >= chartIndex Then
        ws.ChartObjects(chartIndex).Delete
    End If

    DeleteChart = True

    Exit Function

ErrorHandler:
    DeleteChart = False
End Function

'==============================================================================
' CreateSheet - Create a new worksheet
'==============================================================================
Public Function CreateSheet(sheetName As String) As Boolean
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))
    ws.Name = sheetName

    CreateSheet = True

    Exit Function

ErrorHandler:
    CreateSheet = False
End Function

'==============================================================================
' DeleteSheet - Delete a worksheet
'==============================================================================
Public Function DeleteSheet(sheetName As String) As Boolean
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Sheets(sheetName)

    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True

    DeleteSheet = True

    Exit Function

ErrorHandler:
    Application.DisplayAlerts = True
    DeleteSheet = False
End Function

'==============================================================================
' RenameSheet - Rename a worksheet
'==============================================================================
Public Function RenameSheet(oldName As String, newName As String) As Boolean
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Sheets(oldName)

    ws.Name = newName

    RenameSheet = True

    Exit Function

ErrorHandler:
    RenameSheet = False
End Function

'==============================================================================
' ListSheets - List all worksheets
'==============================================================================
Public Function ListSheets() As String
    Dim result As String
    Dim ws As Worksheet

    For Each ws In ActiveWorkbook.Sheets
        result = result & ws.Name & vbCrLf
    Next ws

    ListSheets = result
End Function

'==============================================================================
' ExportToPDF - Export worksheet to PDF
'==============================================================================
Public Function ExportToPDF(sheetName As String, filePath As String) As Boolean
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Sheets(sheetName)

    ws.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=filePath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False

    ExportToPDF = True

    Exit Function

ErrorHandler:
    ExportToPDF = False
End Function

'==============================================================================
' ExportToExcel - Export worksheet to new Excel file
'==============================================================================
Public Function ExportToExcel(sheetName As String, filePath As String) As Boolean
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Sheets(sheetName)

    ws.Copy

    Dim newWB As Workbook
    Set newWB = ActiveWorkbook

    Application.DisplayAlerts = False
    newWB.SaveAs filePath, xlOpenXMLWorkbook
    Application.DisplayAlerts = True

    newWB.Close

    ExportToExcel = True

    Exit Function

ErrorHandler:
    Application.DisplayAlerts = True
    ExportToExcel = False
End Function

'==============================================================================
' SetColumnWidth - Set column width
'==============================================================================
Public Function SetColumnWidth(sheetName As String, columnAddress As String, width As Double) As Boolean
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Sheets(sheetName)

    ws.Columns(columnAddress).ColumnWidth = width

    SetColumnWidth = True

    Exit Function

ErrorHandler:
    SetColumnWidth = False
End Function

'==============================================================================
' SetRowHeight - Set row height
'==============================================================================
Public Function SetRowHeight(sheetName As String, rowAddress As String, height As Double) As Boolean
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Sheets(sheetName)

    ws.Rows(rowAddress).RowHeight = height

    SetRowHeight = True

    Exit Function

ErrorHandler:
    SetRowHeight = False
End Function

'==============================================================================
' MergeCells - Merge cells
'==============================================================================
Public Function MergeCells(sheetName As String, rangeAddress As String) As Boolean
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Sheets(sheetName)

    ws.Range(rangeAddress).Merge

    MergeCells = True

    Exit Function

ErrorHandler:
    MergeCells = False
End Function

'==============================================================================
' UnmergeCells - Unmerge cells
'==============================================================================
Public Function UnmergeCells(sheetName As String, rangeAddress As String) As Boolean
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Sheets(sheetName)

    ws.Range(rangeAddress).UnMerge

    UnmergeCells = True

    Exit Function

ErrorHandler:
    UnmergeCells = False
End Function

'==============================================================================
' SetBorder - Set border for a range
'==============================================================================
Public Function SetBorder(sheetName As String, rangeAddress As String, _
                         Optional borderStyle As Long = xlContinuous, _
                         Optional borderWidth As Long = 1) As Boolean
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Sheets(sheetName)

    Dim rng As Range
    Set rng = ws.Range(rangeAddress)

    With rng.Borders
        .LineStyle = borderStyle
        .Weight = borderWidth
    End With

    SetBorder = True

    Exit Function

ErrorHandler:
    SetBorder = False
End Function

'==============================================================================
' AutoFitColumns - Auto-fit column widths
'==============================================================================
Public Function AutoFitColumns(sheetName As String, Optional rangeAddress As String = "") As Boolean
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Sheets(sheetName)

    If rangeAddress = "" Then
        ws.Columns.AutoFit
    Else
        ws.Range(rangeAddress).Columns.AutoFit
    End If

    AutoFitColumns = True

    Exit Function

ErrorHandler:
    AutoFitColumns = False
End Function

'==============================================================================
' FindAndReplace - Find and replace text
'==============================================================================
Public Function FindAndReplace(sheetName As String, findText As String, replaceText As String, _
                              Optional rangeAddress As String = "") As Boolean
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Sheets(sheetName)

    Dim rng As Range
    If rangeAddress = "" Then
        Set rng = ws.UsedRange
    Else
        Set rng = ws.Range(rangeAddress)
    End If

    rng.Replace _
        What:=findText, _
        Replacement:=replaceText, _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        MatchCase:=False

    FindAndReplace = True

    Exit Function

ErrorHandler:
    FindAndReplace = False
End Function
