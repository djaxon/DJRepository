Attribute VB_Name = "Module1"
Sub StockTrends()

'Variables to Loop below code through worksheets
Dim ws As Worksheet
Dim ws_1 As Worksheet
Set ws_1 = ActiveSheet
For Each ws In ThisWorkbook.Worksheets
ws.Activate
    
    'Variables for Loop through each worksheet
    Dim Ticker As String
    Dim YrlyChange As Double
    Dim Open_Start As Double
    Dim Close_End As Double
    Dim PercentChange As Double
    Dim StockVol As Double
    Dim Summary_Table_Row As Integer

    'Assumptions
    numrows = Range("A1", Range("A1").End(xlDown)).Rows.Count
    Open_Start = Range("C2").Value
    Cells(2, 9).Value = Range("A2").Value
    Summary_Table_Row = 2
    StockVol = 0
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "% Change"
    Range("L1").Value = "Stock Volume"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("P1").Value = "Ticker "
    Range("Q1").Value = "Value"
        
    
    For i = 2 To numrows + 1
            'if statement to make list of each stock / change in price / etc.
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                Ticker = Cells(i + 1, 1).Value
                Close_End = WorksheetFunction.Max(Cells(i, 6).Value, 0.02)
                YrlyChange = Close_End - Open_Start
                StockVol = StockVol + Cells(i, 7).Value
                
                Range("J" & Summary_Table_Row).Value = YrlyChange
                Range("K" & Summary_Table_Row).Value = YrlyChange / Open_Start
                Range("K" & Summary_Table_Row).Style = "Percent"
                Range("L" & Summary_Table_Row).Value = StockVol
                StockVol = 0
                Summary_Table_Row = Summary_Table_Row + 1
                Range("I" & Summary_Table_Row).Value = Ticker
                Open_Start = WorksheetFunction.Max(Cells(i + 1, 3).Value, 0.01)
            
            ElseIf Cells(i + 1, 1).Value = Cells(i, 1).Value And YrlyChange >= 0 Then
                StockVol = StockVol + Cells(i, 7).Value
                Range("J" & Summary_Table_Row - 1).Interior.ColorIndex = 4
            Else
                StockVol = StockVol + Cells(i, 7).Value
                Range("J" & Summary_Table_Row - 1).Interior.ColorIndex = 3
                Close_End = 0
            End If
               
    Next i
    Range("J1").Interior.ColorIndex = 0
    
    'Variables for secondary max / min table
    Dim MaxNum As Double
    Dim MinNum As Double
    Dim MaxStock_Vol As Double
    Dim Ticker_1 As String
    Dim Ticker_2 As String
    Dim Ticker_3 As String
    Dim numrows_2 As Integer
        
        'Assumptions and logic to pull out specific max / min values
        numrows_2 = Range("k1", Range("k1").End(xlDown)).Rows.Count
        MaxNum = WorksheetFunction.Max(Range("K2:" & "K" & numrows_2))
        MinNum = WorksheetFunction.Min(Range("K2:" & "K" & numrows_2))
        MaxStock_Vol = WorksheetFunction.Max(Range("L2:" & "L" & numrows_2))
        Ticker_1 = Cells(WorksheetFunction.Match(MaxNum, Range("K1:K" & numrows_2), 0), 9).Value
        Ticker_2 = Cells(WorksheetFunction.Match(MinNum, Range("K1:K" & numrows_2), 0), 9).Value
        Ticker_3 = Cells(WorksheetFunction.Match(MaxStock_Vol, Range("L1:L" & numrows_2), 0), 9).Value
          
        'printing
        Range("Q2").Value = MaxNum
        Range("Q3").Value = MinNum
        Range("Q2:Q3").Style = "Percent"
        Range("Q4").Value = MaxStock_Vol
        Range("P2").Value = Ticker_1
        Range("P3").Value = Ticker_2
        Range("P4").Value = Ticker_3
        'Autofit cells
        Range("I:Q").EntireColumn.AutoFit

Next
ws_1.Activate

End Sub

