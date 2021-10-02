Attribute VB_Name = "Module1"
Sub stocks()

    'Declaring variables
    Dim ticker As String
    Dim i As Long
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim percentChange As Double
    Dim yearlyChange As Double
    Dim totalStock As LongLong
    Dim x As Integer
    Dim lastRow As Long
    Dim ws As Worksheet
    
    'counter set to specify where to print the final answers later on
    x = 1
    
    'define the stopping point for the loop
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    'loop to pull in the ticker name, opening and closing values across all worksheets
    For Each ws In Worksheets
        ws.Activate
        For i = 2 To lastRow
            totalStock = totalStock + Cells(i, 7).Value
            If Cells(i, 1).Value <> Cells(i - 1, 1) Then
                openingPrice = Cells(i, 3).Value
            ElseIf Cells(i, 1).Value <> Cells(i + 1, 1) Then
                closingPrice = Cells(i, 6).Value
                ticker = Cells(i, 1).Value
                x = x + 1
                'compute the derived answers
                yearlyChange = closingPrice - openingPrice
                If openingPrice <> 0 Then
                    percentChange = (closingPrice - openingPrice) / openingPrice
                formattedPC = FormatPercent(percentChange)
                'print the answers onto the excel sheet
                Cells(x, 12).Value = totalStock
                Cells(x, 9).Value = ticker
                Cells(x, 10).Value = yearlyChange
                Cells(x, 11).Value = formattedPC
                'reset total stock to 0 in preparation for the next ticker
                totalStock = 0
                'highlight the cells for yearly change depending on whether they are negative or positive
                If Cells(x, 10).Value < 0 Then
                    Cells(x, 10).Interior.ColorIndex = 3
                ElseIf Cells(x, 10).Value > 0 Then
                    Cells(x, 10).Interior.ColorIndex = 4
                End If
                End If
            End If
        Next i
        
        'print the titles onto the excel sheet
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
        
        x = 1
    Next ws
    
End Sub







