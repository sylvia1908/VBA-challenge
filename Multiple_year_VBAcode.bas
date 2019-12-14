Attribute VB_Name = "Module1"

Sub ForEachSheet()

    'loop through each sheet
    For Each ws In Worksheets
        ws.Activate
        main
        Greatest
       
    Next ws
    
End Sub

Sub main()

    'to insure ticker and date in order, sort column a to g by a, b asc
    With ActiveSheet.Sort
        .SortFields.Add Key:=Range("A1"), Order:=xlAscending
         .SortFields.Add Key:=Range("B1"), Order:=xlAscending
         .SetRange Range("A:G")
         .Header = xlYes
         .Apply
    End With
    
    'Fill Header
    Cells(1, 12).Value = "Ticker"
    Cells(1, 13).Value = "Price Change"
    Cells(1, 14).Value = "Price Change Percent"
    Cells(1, 15).Value = "Total Volumn"
    Range(Cells(1, 12), Cells(1, 15)).Font.Bold = True

    'get total row number
    rowN = Cells(Rows.Count, 1).End(xlUp).Row '70926
    
    'Start value
    tickerStart = Cells(2, 1).Value
    openStart = Cells(2, 3).Value
    closeStart = Cells(2, 6).Value
    volumnStart = Cells(2, 7).Value
    
    'loop through each row
    For i = 2 To rowN
        tickerRuntime = Cells(i, 1).Value
        openRuntime = Cells(i, 3).Value
        closeRuntime = Cells(i, 6).Value
        volumnRuntime = Cells(i, 7).Value
        If tickerStart = tickerRuntime Then
            volumnStart = volumnStart + volumnRuntime
        ElseIf tickerStart <> tickerRuntime Then
            'get ticker row
            tickerLast = Cells(Rows.Count, 12).End(xlUp).Row + 1
            openEnd = Cells(i - 1, 3).Value
            closeEnd = Cells(i - 1, 6).Value
            volumnStart = volumnStart + volumnRuntime
            'write calculation to cells
            Cells(tickerLast, 12).Value = tickerStart
            'Fill and color price change
            With Cells(tickerLast, 13)
                    .Value = closeEnd - openStart
                    If closeEnd - openStart < 0 Then
                    .Interior.ColorIndex = 3
                
                ElseIf closeEnd - openStartd > 0 Then
                    .Interior.ColorIndex = 4
                Else
                    .Interior.ColorIndex = 7
                End If
            End With
            'Fill PriceChangePercent
            With Cells(tickerLast, 14)
                If openStart <> 0 Then
                    .Value = (closeEnd - openStart) / openStart
                    .NumberFormat = "0.00%"
                End If
            End With
            Cells(tickerLast, 15).Value = volumnStart
        
            'reset Start
            tickerStart = Cells(i, 1).Value
            openStart = Cells(i, 3).Value
            closeStart = Cells(i, 6).Value
            volumnStart = Cells(i, 7).Value
        End If
        Next i

    'Fit
    Columns("A:O").AutoFit
End Sub

Sub Greatest()
    gIncrease = WorksheetFunction.Max(Range("N:N"))
    gDecrease = WorksheetFunction.Min(Range("N:N"))
    gVolumn = WorksheetFunction.Max(Range("O:O"))
    lastrow = Cells(Rows.Count, 14).End(xlUp).Row
    'Fill Header
    Range("Q2").Value = "reatest % increase"
    Range("Q3").Value = "Greatest % Decrease"
    Range("Q4").Value = "Greatest total volume"
    Range("R1").Value = "Ticker"
    Range("S1").Value = "Value"
    
    With Range("Q2:Q4")
        .Interior.Color = RGB(153, 153, 153)
        .Font.Color = vbWhite
        .Font.Bold = True
    End With
    
    With Range("R1:S1")
        .Interior.Color = RGB(153, 153, 153)
        .Font.Color = vbWhite
        .Font.Bold = True
    End With
        
    'FillValue
    Range("S2").Value = gIncrease
    Range("S3").Value = gDecrease
    Range("S4").Value = gVolumn
    'Loop to find each one
    For i = 1 To lastrow
        ticker = Cells(i, 12).Value
        Select Case Cells(i, 14).Value
            Case gIncrease
                Range("R2").Value = ticker
            Case gDecrease
                Range("R3").Value = ticker
        End Select
        If Cells(i, 15).Value = gVolumn Then Range("R4").Value = ticker
     Next i
     
     Columns("Q:S").AutoFit
     Range("S2:S3").NumberFormat = "0.00%"
End Sub


'reset each sheet
Sub Reset()
    For Each ws In Worksheets
    
       ws.Range("L:S").Clear
            
    Next ws
End Sub

