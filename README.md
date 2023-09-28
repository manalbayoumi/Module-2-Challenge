# Module-2-Challenge
Please find here-below my coding for multiyear stock:
Sub StockAnalysis()
    Dim Ticker As String
    Dim OpeningPrice As Double
    Dim ClosingPrice As Double
    
    Dim YearlyChange As Double
    
    Dim PercentageChange As Double
    
    Dim TotalVolume As Double
    Dim LastRow As Long
    Dim Counter As Long
    Dim Counter_Result As Long
    Dim ws As Worksheet
    Dim MaxIncreaseTicker As String
    Dim MaxDecreaseTicker As String
    Dim MaxVolumeTicker As String
    Dim MaxIncrease As Double
    Dim MaxDecrease As Double
    Dim MaxVolume As Double
    
    Dim Sheet As Worksheet
    
    'Set ws = ActiveWorkbook.ActiveSheet
    For Each Sheet In Worksheets
        Sheet.Select
        Set ws = Sheet
        LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        Counter = 2
        Ticker = ws.Cells(Counter, 1)
        OpeningPrice = ws.Cells(Counter, 3)
        TotalVolume = 0
        Counter_Result = 2
        ws.Range("I:Q").Clear
        ws.Cells(1, 9) = "Ticker"
        ws.Cells(1, 10) = "Yearly Change"
        ws.Cells(1, 11) = "Percent Change"
        ws.Cells(1, 12) = "Total Stock Volume"
        Do While Counter <= LastRow
        If ws.Cells(Counter, 1) <> Ticker Then
            ws.Cells(Counter_Result, 9) = Ticker
            ClosingPrice = ws.Cells(Counter - 1, 6)
            ws.Cells(Counter_Result, 10) = ClosingPrice - OpeningPrice
            If ws.Cells(Counter_Result, 10) < 0 Then
                ws.Cells(Counter_Result, 10).Interior.Color = vbRed
            Else
                ws.Cells(Counter_Result, 10).Interior.Color = vbGreen
            End If
            If ClosingPrice = 0 Then
                If OpeningPrice = 0 Then
                    ws.Cells(Counter_Result, 11) = FormatPercent(0, 2)
                Else
                    ws.Cells(Counter_Result, 11) = FormatPercent(-1, 2)
                End If
            Else
                ws.Cells(Counter_Result, 11) = FormatPercent(ws.Cells(Counter_Result, 10) / OpeningPrice, 2)
            End If
            ws.Cells(Counter_Result, 12) = TotalVolume
            If ws.Cells(Counter_Result, 11) < 0 Then
                If ws.Cells(Counter_Result, 11) < MaxDecrease Then
                    MaxDecrease = ws.Cells(Counter_Result, 11)
                    MaxDecreaseTicker = Ticker
                End If
            Else
                If ws.Cells(Counter_Result, 11) > MaxIncrease Then
                    MaxIncrease = ws.Cells(Counter_Result, 11)
                    MaxIncreaseTicker = Ticker
                End If
            End If
            If ws.Cells(Counter_Result, 12) > MaxVolume Then
                MaxVolume = ws.Cells(Counter_Result, 12)
                MaxVolumeTicker = Ticker
            End If
            Ticker = ws.Cells(Counter, 1)
            TotalVolume = ws.Cells(Counter, 7)
            OpeningPrice = ws.Cells(Counter, 3)
            Counter_Result = Counter_Result + 1
        Else
            TotalVolume = TotalVolume + ws.Cells(Counter, 7)
        End If
        Counter = Counter + 1
       Loop
       ClosingPrice = ws.Cells(Counter - 1, 6)
       ws.Cells(Counter_Result, 10) = ClosingPrice - OpeningPrice
        If ws.Cells(Counter_Result, 10) < 0 Then
            ws.Cells(Counter_Result, 10).Interior.Color = vbRed
        Else
            ws.Cells(Counter_Result, 10).Interior.Color = vbGreen
        End If
        If ClosingPrice = 0 Then
            If OpeningPrice = 0 Then
                ws.Cells(Counter_Result, 11) = FormatPercent(0, 2)
            Else
                ws.Cells(Counter_Result, 11) = FormatPercent(-1, 2)
            End If
        Else
            ws.Cells(Counter_Result, 11) = FormatPercent(ws.Cells(Counter_Result, 10) / ClosingPrice, 2)
        End If
        ws.Cells(Counter_Result, 9) = ws.Cells(Counter - 1, 1)
        ws.Cells(Counter_Result, 12) = TotalVolume
        If ws.Cells(Counter_Result, 11) < 0 Then
            If ws.Cells(Counter_Result, 11) < MaxDecrease Then
                MaxDecrease = ws.Cells(Counter_Result, 11)
                MaxDecreaseTicker = ws.Cells(Counter_Result, 9)
            End If
        Else
            If ws.Cells(Counter_Result, 11) > MaxIncrease Then
                MaxIncrease = ws.Cells(Counter_Result, 11)
                MaxIncreaseTicker = ws.Cells(Counter_Result, 9)
            End If
        End If
        If ws.Cells(Counter_Result, 12) > MaxVolume Then
            MaxVolume = ws.Cells(Counter_Result, 12)
            MaxVolumeTicker = ws.Cells(Counter_Result, 9)
        End If
        ws.Range("O2") = "Greatest % Increase"
        ws.Range("p1") = "Ticker"
        ws.Range("q1") = "Value"
        ws.Range("p2") = MaxIncreaseTicker
        ws.Range("q2") = FormatPercent(MaxIncrease, 2)
        
        ws.Range("O3") = "Greatest % Decrease"
        ws.Range("p3") = MaxDecreaseTicker
        ws.Range("q3") = FormatPercent(MaxDecrease, 2)
        
        ws.Range("O4") = "Greatest Total Volume"
        ws.Range("p4") = MaxVolumeTicker
        ws.Range("q4") = MaxVolume
        
    Next
    MsgBox ("ALL DONE")
      
       
       
End Sub

Please find my coding for formating for excel file:

![image](https://github.com/manalbayoumi/Module-2-Challenge/assets/139724159/60aef50f-53ad-44c9-b984-bb09db5a52a9)


Coding for alphabetical testing file:
Sub StockAnalysis()
    Dim Ticker As String
    Dim OpeningPrice As Double
    Dim ClosingPrice As Double
    
    Dim YearlyChange As Double
    
    Dim PercentageChange As Double
    
    Dim TotalVolume As Double
    Dim LastRow As Long
    Dim Counter As Long
    Dim Counter_Result As Long
    Dim ws As Worksheet
    Dim MaxIncreaseTicker As String
    Dim MaxDecreaseTicker As String
    Dim MaxVolumeTicker As String
    Dim MaxIncrease As Double
    Dim MaxDecrease As Double
    Dim MaxVolume As Double
    
    Dim Sheet As Worksheet
    
    'Set ws = ActiveWorkbook.ActiveSheet
    For Each Sheet In Worksheets
        Sheet.Select
        Set ws = Sheet
        LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        Counter = 2
        Ticker = ws.Cells(Counter, 1)
        OpeningPrice = ws.Cells(Counter, 3)
        TotalVolume = 0
        Counter_Result = 2
        ws.Range("I:Q").Clear
        ws.Cells(1, 9) = "Ticker"
        ws.Cells(1, 10) = "Yearly Change"
        ws.Cells(1, 11) = "Percent Change"
        ws.Cells(1, 12) = "Total Stock Volume"
        Do While Counter <= LastRow
        If ws.Cells(Counter, 1) <> Ticker Then
            ws.Cells(Counter_Result, 9) = Ticker
            ClosingPrice = ws.Cells(Counter - 1, 6)
            ws.Cells(Counter_Result, 10) = ClosingPrice - OpeningPrice
            If ws.Cells(Counter_Result, 10) < 0 Then
                ws.Cells(Counter_Result, 10).Interior.Color = vbRed
            Else
                ws.Cells(Counter_Result, 10).Interior.Color = vbGreen
            End If
            If ClosingPrice = 0 Then
                If OpeningPrice = 0 Then
                    ws.Cells(Counter_Result, 11) = FormatPercent(0, 2)
                Else
                    ws.Cells(Counter_Result, 11) = FormatPercent(-1, 2)
                End If
            Else
                ws.Cells(Counter_Result, 11) = FormatPercent(ws.Cells(Counter_Result, 10) / OpeningPrice, 2)
            End If
            ws.Cells(Counter_Result, 12) = TotalVolume
            If ws.Cells(Counter_Result, 11) < 0 Then
                If ws.Cells(Counter_Result, 11) < MaxDecrease Then
                    MaxDecrease = ws.Cells(Counter_Result, 11)
                    MaxDecreaseTicker = Ticker
                End If
            Else
                If ws.Cells(Counter_Result, 11) > MaxIncrease Then
                    MaxIncrease = ws.Cells(Counter_Result, 11)
                    MaxIncreaseTicker = Ticker
                End If
            End If
            If ws.Cells(Counter_Result, 12) > MaxVolume Then
                MaxVolume = ws.Cells(Counter_Result, 12)
                MaxVolumeTicker = Ticker
            End If
            Ticker = ws.Cells(Counter, 1)
            TotalVolume = ws.Cells(Counter, 7)
            OpeningPrice = ws.Cells(Counter, 3)
            Counter_Result = Counter_Result + 1
        Else
            TotalVolume = TotalVolume + ws.Cells(Counter, 7)
        End If
        Counter = Counter + 1
       Loop
       ClosingPrice = ws.Cells(Counter - 1, 6)
       ws.Cells(Counter_Result, 10) = ClosingPrice - OpeningPrice
        If ws.Cells(Counter_Result, 10) < 0 Then
            ws.Cells(Counter_Result, 10).Interior.Color = vbRed
        Else
            ws.Cells(Counter_Result, 10).Interior.Color = vbGreen
        End If
        If ClosingPrice = 0 Then
            If OpeningPrice = 0 Then
                ws.Cells(Counter_Result, 11) = FormatPercent(0, 2)
            Else
                ws.Cells(Counter_Result, 11) = FormatPercent(-1, 2)
            End If
        Else
            ws.Cells(Counter_Result, 11) = FormatPercent(ws.Cells(Counter_Result, 10) / ClosingPrice, 2)
        End If
        ws.Cells(Counter_Result, 9) = ws.Cells(Counter - 1, 1)
        ws.Cells(Counter_Result, 12) = TotalVolume
        If ws.Cells(Counter_Result, 11) < 0 Then
            If ws.Cells(Counter_Result, 11) < MaxDecrease Then
                MaxDecrease = ws.Cells(Counter_Result, 11)
                MaxDecreaseTicker = ws.Cells(Counter_Result, 9)
            End If
        Else
            If ws.Cells(Counter_Result, 11) > MaxIncrease Then
                MaxIncrease = ws.Cells(Counter_Result, 11)
                MaxIncreaseTicker = ws.Cells(Counter_Result, 9)
            End If
        End If
        If ws.Cells(Counter_Result, 12) > MaxVolume Then
            MaxVolume = ws.Cells(Counter_Result, 12)
            MaxVolumeTicker = ws.Cells(Counter_Result, 9)
        End If
        ws.Range("O2") = "Greatest % Increase"
        ws.Range("p1") = "Ticker"
        ws.Range("q1") = "Value"
        ws.Range("p2") = MaxIncreaseTicker
        ws.Range("q2") = FormatPercent(MaxIncrease, 2)
        
        ws.Range("O3") = "Greatest % Decrease"
        ws.Range("p3") = MaxDecreaseTicker
        ws.Range("q3") = FormatPercent(MaxDecrease, 2)
        
        ws.Range("O4") = "Greatest Total Volume"
        ws.Range("p4") = MaxVolumeTicker
        ws.Range("q4") = MaxVolume
        
Next
      
       
       
End Sub


