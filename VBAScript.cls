VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Stock_market()
    Dim ws As Worksheet
    For Each ws In Worksheets
        ' Assuming data starts from row 2 in columns A to G
        Dim LastRow As Long
        LastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        ' Initialize headers
        ws.Range("I1:L1").Value = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume")
        
        ' Initialize variables
        Dim Ticker As String
        Dim OpenPrice As Double
        Dim ClosePrice As Double
        Dim YearlyChange As Double
        Dim PercentChange As Double
        Dim TotalVolume As Double
        Dim OutputRow As Long
        
        ' Initialize starting values
        Ticker = ws.Cells(2, 1).Value
        OpenPrice = ws.Cells(2, 3).Value
        TotalVolume = 0
        OutputRow = 2
        
        ' Loop through rows
        For i = 2 To LastRow
            If ws.Cells(i, 1).Value <> Ticker Then
                ' Output results for previous Ticker
                OutputResults ws, OutputRow, Ticker, OpenPrice, ClosePrice, YearlyChange, PercentChange, TotalVolume
                
                ' Reset variables for the new Ticker
                Ticker = ws.Cells(i, 1).Value
                OpenPrice = ws.Cells(i, 3).Value
                TotalVolume = 0
                OutputRow = OutputRow + 1
            End If
            
            ' Accumulate stock volume
            TotalVolume = TotalVolume + ws.Cells(i, 7).Value
            
            ' Update close price for the Ticker
            ClosePrice = ws.Cells(i, 6).Value
        Next i
        
        ' Output results for the last Ticker
        OutputResults ws, OutputRow, Ticker, OpenPrice, ClosePrice, YearlyChange, PercentChange, TotalVolume
    Next ws
End Sub

Sub OutputResults(ws As Worksheet, OutputRow As Long, Ticker As String, OpenPrice As Double, ClosePrice As Double, ByRef YearlyChange As Double, ByRef PercentChange As Double, TotalVolume As Double)
    ' Calculate Yearly Change and Percent Change
    YearlyChange = ClosePrice - OpenPrice
    If OpenPrice <> 0 Then
        PercentChange = YearlyChange / OpenPrice
    Else
        PercentChange = 0
    End If
    
    ' Output results to the worksheet
    ws.Cells(OutputRow, 9).Value = Ticker
    ws.Cells(OutputRow, 10).Value = YearlyChange
    ws.Cells(OutputRow, 11).Value = PercentChange
    ws.Cells(OutputRow, 12).Value = TotalVolume
    
    ' Apply formatting
    FormatResults ws, OutputRow
End Sub

Sub FormatResults(ws As Worksheet, OutputRow As Long)
    ' Apply formatting for Yearly Change
    If ws.Cells(OutputRow, 10).Value > 0 Then
        ws.Cells(OutputRow, 10).Interior.ColorIndex = 4 'Green
    ElseIf ws.Cells(OutputRow, 10).Value < 0 Then
        ws.Cells(OutputRow, 10).Interior.ColorIndex = 3 'Red
    End If
    
    ' Apply formatting for Percent Change
    ws.Cells(OutputRow, 11).NumberFormat = "0.00%"
    If ws.Cells(OutputRow, 11).Value > 0 Then
        ws.Cells(OutputRow, 11).Interior.ColorIndex = 4 'Green
    ElseIf ws.Cells(OutputRow, 11).Value < 0 Then
        ws.Cells(OutputRow, 11).Interior.ColorIndex = 3 'Red
    End If
End Sub

