Attribute VB_Name = "Module1"
Sub stock_analysis():

    ' Declare variables
    Dim total As Double
    Dim i As Long
    Dim j As Long
    Dim change As Single
    Dim start As Long
    Dim rowCount As Long
    Dim days As Integer
    Dim percentChange As Double
    Dim dailyChange As Single
    Dim averageChange As Single
    Dim ws As Worksheet

    For Each ws In Worksheets
        ' Set initial values for each worksheet
        j = 0
        total = 0
        change = 0
        start = 2
        dailyChange = 0

        ' Print summary column title for each worksheet
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"

        ' Get the row number of the last row with data
        rowCount = Cells(Rows.Count, "A").End(xlUp).Row

        For i = 2 To rowCount

            ' If ticker changes then print results
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                ' Stores total volume in total
                total = total + ws.Cells(i, 7).Value

                ' Handle zero total volume
                If total = 0 Then
                    ' print the results
                    ws.Range("I" & 2 + j).Value = Cells(i, 1).Value
                    ws.Range("J" & 2 + j).Value = 0
                    ws.Range("K" & 2 + j).Value = "%" & 0
                    ws.Range("L" & 2 + j).Value = 0

                Else
                    ' Find first non zero starting value
                    If ws.Cells(start, 3) = 0 Then
                        For find_value = start To i
                            If ws.Cells(find_value, 3).Value <> 0 Then
                                start = find_value
                                Exit For
                            End If
                        Next find_value
                    End If

                    ' Calculate Change
                    change = (ws.Cells(i, 6) - ws.Cells(start, 3))
                    percentChange = Round((change / ws.Cells(start, 3) * 100), 2)

                    ' Next stock ticker
                    start = i + 1

                    ' Print results to a separate worksheet
                    ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
                    ws.Range("J" & 2 + j).Value = Round(change, 2)
                    ws.Range("K" & 2 + j).Value = "%" & percentChange
                    ws.Range("L" & 2 + j).Value = total
                    
                    ' Change fill cell colors
                    If change > 0 Then
                        ' color positive green
                        ws.Range("J" & 2 + j).Interior.ColorIndex = 4
                    ElseIf change < 0 Then
                        ' color negatives red
                        ws.Range("J" & 2 + j).Interior.ColorIndex = 3
                    Else
                        ws.Range("J" & 2 + j).Interior.ColorIndex = 0
                    End If

                End If

                ' Reset variables for new stock ticker
                total = 0
                change = 0
                j = j + 1
                days = 0
                dailyChange = 0

            ' If ticker is still the same add total volume
            Else
                total = total + ws.Cells(i, 7).Value

            End If

        Next i

    Next ws

End Sub


