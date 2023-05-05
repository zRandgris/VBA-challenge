Attribute VB_Name = "Module2"
Sub Module2_Loop():

    For Each ws In Worksheets
        Dim LastRow As Long
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
            
        'MsgBox ("Last Row is :" & LastRow)
        'Creating Headers for the summary table.
        ws.Cells(1, 9) = "Ticker"
        ws.Cells(1, 10) = "Yearly Change"
        ws.Cells(1, 11) = "Percent Change"
        ws.Cells(1, 12) = "Total Stock Volume"
        ' Autofit to colum for better view
        
        ws.Columns("I:M").AutoFit
        ws.Columns("K").NumberFormat = "0.00%"
        
        ' Making Variable
        Dim Tickerid As String
        Dim TickerOpen As Double
        Dim TickerClose As Double
        Dim TickerVol As Double
        Dim TickerChange As Currency
        Dim TickerPChange As Double
        Dim SummaryRow As Integer
        Dim TickerVolS As String
        
        SummaryRow = 2
        TickerVol = 0
        TickerOpen = 0
        TickerClose = 0
        TickerChange = 0
        TickerPChange = 0
        
        ' Returning value while checking Ticker Name stayed the same.
        TickerOpen = Cells(2, 3).Value
        For i = 2 To LastRow
            'MsgBox (TickerOpen)
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ' Storing Ticker's Name into Var
                Tickerid = ws.Cells(i, 1).Value
                
                'Total Stock Volume
                
                TickerVol = TickerVol + ws.Cells(i, 7).Value
                'MsgBox ("Vol:" & TickerVol)
                'TickerVolS = TickerVol
                'Ticker Closing Value
                
                TickerClose = ws.Cells(i, 6).Value
                
                'MsgBox ("OPEN:" & TickerOpen)
                'MsgBox ("Close: " & TickerClose)
                
                'Calculating Ticker Change
                
                TickerChange = TickerClose - TickerOpen
                                
                'Calculating % Change
                
                TickerPChange = TickerChange / TickerOpen
                
                
                TickerPChangeF = Format(TickerPChange, "0.00%")
                                                
                'Printing Ticker= I, Year Change = J , Percent Change = K, Total= L
                
                ws.Range("I" & SummaryRow).Value = Tickerid
                ws.Range("J" & SummaryRow).Value = TickerChange
                ws.Range("K" & SummaryRow).Value = TickerPChangeF
                ws.Range("L" & SummaryRow).Value = TickerVol
                SummaryRow = SummaryRow + 1
                
                TickerVol = 0
                TickerOpen = ws.Cells(i + 1, 3).Value
                
            Else
                'Total Stock Volume
                
                'MsgBox ("I am Row#:" & i)
                'MsgBox ("Currently Have:" & TickerVol)
                TickerVol = TickerVol + ws.Cells(i, 7).Value
               
                
                
                'Ticker Opening Value
               
                
            End If

            
        Next i
        
        ' Condition Format
        Dim LastRowS As Long
        LastRowS = ws.Cells(Rows.Count, 11).End(xlUp).Row
        For x = 2 To LastRowS
            
            If ws.Cells(x, 11).Value >= 0 Then
                ws.Cells(x, 11).Interior.ColorIndex = 4
                ws.Cells(x, 10).Interior.ColorIndex = 4
            ElseIf ws.Cells(x, 11).Value < 0 Then
                ws.Cells(x, 11).Interior.ColorIndex = 3
                ws.Cells(x, 10).Interior.ColorIndex = 3
            End If
        Next x
           
        ' Adding Fuctionality
        Dim MaxP As Double
        Dim MinP As Double
        Dim MaxV As Double
        
        
        'rg = Range(K2, K91)
        MaxP = WorksheetFunction.Max(ws.Range("K2: K" & LastRowS))
        'MsgBox ("Max % is" & Maxp)
        MaxPF = Format(MaxP, "0.00%")
        ws.Range("Q2") = MaxPF
        ws.Range("O2") = "Greatest%Increase"
        '---------Greatest%---------------------------------
        MinP = WorksheetFunction.Min(ws.Range("K2: K" & LastRowS))
        MinPF = Format(MinP, "0.00%")
        ws.Range("Q3") = MinPF
        ws.Range("O3") = "Greatest%Decrease"
        '----------GreatestDecrease---------------------------------
        MaxV = WorksheetFunction.Max(ws.Range("L2: L" & LastRowS))
        ws.Range("Q4") = MaxV
        ws.Range("O4") = "GreatestTotalVolume"
        '------------GreatesTotalVolume--------------------------------
        'For y = 2 To LastRowS
        ws.Range("P1") = "Ticker"
        ws.Range("Q1") = "Value"
        '---------------------------------------------------------
        
        ' Matching the value and returning the Ticker
        For y = 2 To LastRowS
            If ws.Cells(y, 11) = MaxP Then
            ws.Range("P2") = ws.Cells(y, 9)
            
            ElseIf ws.Cells(y, 11) = MinP Then
            ws.Range("P3") = ws.Cells(y, 9)
            
            ElseIf ws.Cells(y, 12) = MaxV Then
            ws.Range("P4") = ws.Cells(y, 9)
            End If
        Next y
      ws.Columns("O:Q").AutoFit
               
            

    Next ws


End Sub

