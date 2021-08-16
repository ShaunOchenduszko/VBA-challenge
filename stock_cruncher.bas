Attribute VB_Name = "stock_cruncher"
Sub stockcrunch():

'loop for all sheets
For Each ws In Worksheets

    'set summary table header
    ws.Range("I1").Value = "Ticker Symbol"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percentage Change"
    ws.Range("L1").Value = "Total Stock Volume"

    'resize summary table columns
    ws.Range("I1:L1").Columns.AutoFit

    'set variables
    Dim ssymbol As String
    Dim totalvolume As Double
    Dim summarytblrow As Integer
    Dim opener As Double
    Dim ender As Double
    Dim yearchange As Double
    Dim percentchange As Double
    Dim lastRow As Long

    'set last row of data
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    'set summary table row
    summarytblrow = 2

    'set initial opening price
    opener = ws.Cells(2, 3).Value

    'begin data check and aggregation loop
    For i = 2 To lastRow
    
        'if next cell doesn't match current
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            'set ticker symbol
            ssymbol = ws.Cells(i, 1).Value
            totalvolume = totalvolume + ws.Cells(i, 7).Value
            'set ending price
            ender = ws.Cells(i, 6).Value
            'set year change price
            yearchange = ender - opener
        
            'print the symbol to the summary table
            ws.Range("I" & summarytblrow).Value = ssymbol

                'print the yearchange to the summary table
                If yearchange > 0 Then
                    'color the yearchange in the summary table
                    ws.Range("J" & summarytblrow).Value = yearchange
                    ws.Range("J" & summarytblrow).Interior.ColorIndex = 4
                    Else
                    ws.Range("J" & summarytblrow).Value = yearchange
                    ws.Range("J" & summarytblrow).Interior.ColorIndex = 3
                End If
        
            'print the percentage change to the summary table
            'set percentage change to 0 if year change is 0
            If yearchange = 0 Then
                ws.Range("K" & summarytblrow).Value = "0"
                ws.Range("K" & summarytblrow).NumberFormat = "0.00%"
            
            'set percentage change to null if opener if 0
            ElseIf opener = 0 Then
                ws.Range("K" & summarytblrow).Value = "Null"
            
            'otherwise calculate percentage change
            Else
                yearchange = yearchange / opener
                ws.Range("K" & summarytblrow).Value = yearchange
                ws.Range("K" & summarytblrow).NumberFormat = "0.00%"
            End If
            
            'print the total volume to the summary table
            ws.Range("L" & summarytblrow).Value = totalvolume
        
            'add one rowto the summary table
            summarytblrow = summarytblrow + 1
        
            'reset new stock begining price
            opener = ws.Cells(i + 1, 3).Value
      
            'reset the total volume
            totalvolume = 0

            'if the next cell is the same symbol
            Else

        'add tototal volume
        totalvolume = totalvolume + ws.Cells(i, 7).Value

        End If

    Next i

    'find/set last row of summary table
    Dim sumlastrow As Long
    sumlastrow = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row
        
    'set bonus table
    ws.Range("N2").Value = "Greatest Percentage Increase"
    ws.Range("P2").Interior.ColorIndex = 4
    ws.Range("N3").Value = "Greatest Percentage Decrease"
    ws.Range("P3").Interior.ColorIndex = 3
    ws.Range("N4").Value = "Greatest Total Volume"
    ws.Range("O1").Value = "Ticker Symbol"
    ws.Range("P1").Value = "Value"
    ws.Range("P2,P3").NumberFormat = "0.00%"
    
        'loop for greatinc column
        Dim greatest As Double
        greatest = 1
        For greatinc = 2 To sumlastrow
        
            If ws.Cells(greatinc, 11).Value > greatest Then
            greatest = ws.Cells(greatinc, 11)
            ws.Range("P2").Value = ws.Cells(greatinc, 11)
            ws.Range("O2").Value = ws.Cells(greatinc, 9)
            End If
                    
        Next greatinc
        
        'loop for greatdec column
        Dim least As Double
        least = 1
        For greatdec = 2 To sumlastrow
        
            If ws.Cells(greatdec, 11).Value < least Then
            least = ws.Cells(greatdec, 11)
            ws.Range("P3").Value = ws.Cells(greatdec, 11)
            ws.Range("O3").Value = ws.Cells(greatdec, 9)
            End If
                            
        Next greatdec
        
        'loop for totvol column
        Dim totvol As Double
        totvol = 1
        For bestvolume = 2 To sumlastrow
        
            If ws.Cells(bestvolume, 12).Value > totvol Then
            totvol = ws.Cells(bestvolume, 12)
            ws.Range("P4").Value = ws.Cells(bestvolume, 12)
            ws.Range("O4").Value = ws.Cells(bestvolume, 9)
            End If
                    
        Next bestvolume

    'resize bonus table columns
    ws.Range("N1:P4").Columns.AutoFit

Next ws

MsgBox ("Complete")

End Sub
