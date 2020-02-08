Attribute VB_Name = "mBusinessLogic"
Option Explicit

' Stock Data Columns
Const COL_TICKER = 1
Const COL_DATE = 2
Const COL_OPEN = 3
Const COL_HIGH = 4
Const COL_LOW = 5
Const COL_CLOSE = 6
Const COL_VOL = 7

' Report Columns
Const RPT_TICKER = 1
Const RPT_CHANGE = 2
Const RPT_CHANGE_PCT = 3
Const RPT_VOLUME = 4
Const RPT_NOTEABLE = 6

' Notable Stock Indices
Const GAIN = 1
Const LOSS = 2
Const VOLUME = 3

Public Sub Initialize()
    Dim fso As Scripting.FileSystemObject
    Dim folderStockData As Scripting.folder
    Dim fileStockData As Scripting.File
    Dim iRow As Long
    Dim sFilename As String
    
    Set fso = New Scripting.FileSystemObject
    Set folderStockData = fso.GetFolder(wksConfig.Range("StockDataDirectory"))
    wksRawData.Range("A2:A100").ClearContents
    iRow = 2
    For Each fileStockData In folderStockData.Files
        sFilename = fileStockData.Name
        If Left(sFilename, 1) <> "~" Then
            wksRawData.Cells(iRow, 1) = fileStockData.Name
            iRow = iRow + 1
        End If
    Next fileStockData
End Sub

Public Sub AnalyzeStocks()
    Dim wkbStockData As Workbook
    Dim sStockDataDirectory As String
    Dim sStockDataFilename As String
    Dim wksStockData As Worksheet
    Dim wksReport As Worksheet

    Dim iNextReportRow As Long
    
    ' With more effort, I would create a Class named Stock
    Dim NotableStocks(1 To 3, 1 To 4) As Variant
    
    sStockDataDirectory = wksConfig.Range("StockDataDirectory")
    sStockDataFilename = wksRawData.Range("StockDataFileAnchor").Offset(wksConfig.Range("StockDataFileIndex"), 0)
    RemoveReportSheets ThisWorkbook
    Set wkbStockData = OpenWorkbook(sStockDataDirectory, sStockDataFilename)
    'AnalyzeSheet wkbStockData.Sheets(1), wksReport, iNextReportRow, NotableStocks
    For Each wksStockData In wkbStockData.Sheets
        Set wksReport = Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wksReport.Name = "Report_" + wksStockData.Name
        iNextReportRow = 2
        InitializeNotableStocks NotableStocks
        AnalyzeSheet wksStockData, wksReport, iNextReportRow, NotableStocks
        wksReport.Range("G2:J4") = NotableStocks
        FormatReport wksReport
    Next wksStockData
End Sub

Public Sub AnalyzeSheet(wksData As Worksheet, wksReport As Worksheet, ByRef iNextReportRow As Long, ByRef NotableStocks As Variant)
    Dim nLastRow As Long
    Dim iRow As Long
    Dim sTicker As String
    Dim sLastTicker As String
    Dim dInitialValue As Double
    Dim dFinalValue As Double
    Dim dPercentChange As Double
    Dim dCumVolume As Double
    
    If IsEmpty(wksData.Cells(2, COL_TICKER)) Then
        nLastRow = 0
    Else
        nLastRow = wksData.Cells(1, COL_TICKER).End(xlDown).Row
    End If

    sLastTicker = wksData.Cells(2, COL_TICKER)
    dInitialValue = wksData.Cells(2, COL_OPEN)
    For iRow = 2 To nLastRow + 1 'Run into a final blank row, to force output of the last ticker
        sTicker = wksData.Cells(iRow, COL_TICKER)
        dCumVolume = dCumVolume + wksData.Cells(iRow, COL_VOL)
        If sTicker <> sLastTicker Then
            ' Populate Report Row for the Stock
            wksReport.Cells(iNextReportRow, RPT_TICKER) = sLastTicker
            wksReport.Cells(iNextReportRow, RPT_CHANGE) = dFinalValue - dInitialValue
            If dInitialValue = 0 Then
                If dFinalValue = 0 Then
                    dPercentChange = 0
                Else
                    dPercentChange = 100
                End If
            Else
                dPercentChange = (dFinalValue - dInitialValue) / dInitialValue
            End If
            wksReport.Cells(iNextReportRow, RPT_CHANGE_PCT) = dPercentChange
            wksReport.Cells(iNextReportRow, RPT_VOLUME) = dCumVolume
            
            ' Check for Notable Stocks
            If dPercentChange > NotableStocks(GAIN, RPT_CHANGE_PCT) Then
                NotableStocks(GAIN, RPT_TICKER) = sLastTicker
                NotableStocks(GAIN, RPT_CHANGE) = dFinalValue - dInitialValue
                NotableStocks(GAIN, RPT_CHANGE_PCT) = dPercentChange
                NotableStocks(GAIN, RPT_VOLUME) = dCumVolume
            End If
                
            If dPercentChange < NotableStocks(LOSS, RPT_CHANGE_PCT) Then
                NotableStocks(LOSS, RPT_TICKER) = sLastTicker
                NotableStocks(LOSS, RPT_CHANGE) = dFinalValue - dInitialValue
                NotableStocks(LOSS, RPT_CHANGE_PCT) = dPercentChange
                NotableStocks(LOSS, RPT_VOLUME) = dCumVolume
            End If
                
            If dCumVolume > NotableStocks(VOLUME, RPT_VOLUME) Then
                NotableStocks(VOLUME, RPT_TICKER) = sLastTicker
                NotableStocks(VOLUME, RPT_CHANGE) = dFinalValue - dInitialValue
                NotableStocks(VOLUME, RPT_CHANGE_PCT) = dPercentChange
                NotableStocks(VOLUME, RPT_VOLUME) = dCumVolume
            End If
            
            ' Update counters
            sLastTicker = sTicker
            dInitialValue = wksData.Cells(iRow, COL_OPEN)
            dCumVolume = wksData.Cells(iRow, COL_VOL)
            iNextReportRow = iNextReportRow + 1
        End If
        dFinalValue = wksData.Cells(iRow, COL_CLOSE)
    Next iRow
End Sub

Public Sub RemoveReportSheets(wkb As Workbook)
    Dim wks As Worksheet
    
    For Each wks In wkb.Sheets
        If Left(wks.Name, 7) = "Report_" Then
            Application.DisplayAlerts = False
            wks.Delete
            Application.DisplayAlerts = True
        End If
    Next wks
End Sub

Public Sub FormatReport(wksReport As Worksheet)
    With wksReport
        .Activate
        
        .Cells(1, RPT_TICKER) = "Ticker"
        .Cells(1, RPT_CHANGE) = "Gain/Loss"
        .Cells(1, RPT_CHANGE_PCT) = "Return"
        .Cells(1, RPT_VOLUME) = "Volume"
        
        .Cells(1, RPT_NOTEABLE) = "Notable Stocks"
        .Cells(1, RPT_NOTEABLE + RPT_TICKER) = "Ticker"
        .Cells(1, RPT_NOTEABLE + RPT_CHANGE) = "Gain/Loss"
        .Cells(1, RPT_NOTEABLE + RPT_CHANGE_PCT) = "Return"
        .Cells(1, RPT_NOTEABLE + RPT_VOLUME) = "Volume"
        
        .Cells(2, RPT_NOTEABLE) = "Greatest Percent Increase"
        .Cells(3, RPT_NOTEABLE) = "Greatest Percent Decrease"
        .Cells(4, RPT_NOTEABLE) = "Greatest Total Volume"

        .Range("A1:J1").Font.Bold = True
        .Range("A1:J1").HorizontalAlignment = xlCenter
        .Columns(RPT_NOTEABLE).ColumnWidth = 26
    End With

    Dim nLastRow As Long
    If IsEmpty(wksReport.Cells(2, 1)) Then
        nLastRow = 0
    Else
        nLastRow = wksReport.Cells(1, 1).End(xlDown).Row
    End If
    nLastRow = nLastRow
    
    wksReport.Range(wksReport.Cells(2, RPT_CHANGE), wksReport.Cells(nLastRow, RPT_CHANGE)).NumberFormat = "$#,##0.00"
    wksReport.Range(wksReport.Cells(2, RPT_CHANGE_PCT), wksReport.Cells(nLastRow, RPT_CHANGE_PCT)).NumberFormat = "0.00%"
    wksReport.Range(wksReport.Cells(2, RPT_VOLUME), wksReport.Cells(nLastRow, RPT_VOLUME)).NumberFormat = "#,##0"
    
    wksReport.Range(wksReport.Cells(2, RPT_NOTEABLE + RPT_CHANGE), wksReport.Cells(4, RPT_NOTEABLE + RPT_CHANGE)).NumberFormat = "$#,##0.00"
    wksReport.Range(wksReport.Cells(2, RPT_NOTEABLE + RPT_CHANGE_PCT), wksReport.Cells(4, RPT_NOTEABLE + RPT_CHANGE_PCT)).NumberFormat = "0.00%"
    wksReport.Range(wksReport.Cells(2, RPT_NOTEABLE + RPT_VOLUME), wksReport.Cells(4, RPT_NOTEABLE + RPT_VOLUME)).NumberFormat = "#,##0"
    
    wksReport.UsedRange.FormatConditions.Delete
    wksReport.Range(wksReport.Cells(2, RPT_CHANGE), wksReport.Cells(nLastRow, RPT_CHANGE_PCT)).Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0"
    With Selection.FormatConditions(1)
        .Interior.PatternColorIndex = xlAutomatic
        .Interior.Color = 13434828
        .Interior.TintAndShade = 0
        .StopIfTrue = False
    End With

    wksReport.Range(wksReport.Cells(2, RPT_CHANGE), wksReport.Cells(nLastRow, RPT_CHANGE_PCT)).Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
    With Selection.FormatConditions(2)
        .Interior.PatternColorIndex = xlAutomatic
        .Interior.Color = 13421823
        .Interior.TintAndShade = 0
        .StopIfTrue = False
    End With
    wksReport.Range("A1").Select
End Sub

Public Sub InitializeNotableStocks(ByRef NotableStocks() As Variant)
    Dim i As Integer
    
    For i = GAIN To VOLUME
        NotableStocks(i, RPT_TICKER) = "Dummy"
        NotableStocks(i, RPT_CHANGE) = 0#
        NotableStocks(i, RPT_CHANGE_PCT) = 0#
        NotableStocks(i, RPT_VOLUME) = 0
    Next i
End Sub


