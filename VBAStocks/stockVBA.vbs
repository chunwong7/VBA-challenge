Attribute VB_Name = "Module1"
Sub stocks()
        Dim lastrow As Double
        Dim lastcolumn As Double
        Dim ticker As String
        Dim volume As Double
        Dim opening As Double
        Dim closing As Double
        Dim yearlychange As Double
        Dim yearlypercent As Double
        Dim newrow As Double
        Dim ws As Worksheet
        For Each ws In Worksheets
            lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
            lastcolumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column
            'MsgBox (lastrow)
            'MsgBox (lastcolumn)
            opening = ws.Cells(2, 3)
            newrow = 2
            ws.Cells(1, 9) = "Ticker"
            ws.Cells(1, 10) = "Yearly Change"
            ws.Cells(1, 11) = "Percent Change"
            ws.Cells(1, 12) = "Total Stock Volume"
            For i = 2 To lastrow
                volume = volume + ws.Cells(i, 7)
                If ws.Cells(i, 1) <> ws.Cells(i + 1, 1) Then
                    ticker = ws.Cells(i, 1)
                    closing = ws.Cells(i, 6)
                    yearlychange = closing - opening
                    yearlypercent = yearlychange / opening
                    ws.Cells(newrow, 9) = ticker
                    ws.Cells(newrow, 10) = yearlychange
                    If yearlychange > 0 Then
                        ws.Cells(newrow, 10).Interior.ColorIndex = 4
                    ElseIf yearlychange < 0 Then
                        ws.Cells(newrow, 10).Interior.ColorIndex = 3
                    End If
                    ws.Cells(newrow, 11).NumberFormat = "0.00%"
                    ws.Cells(newrow, 11) = yearlypercent
                    ws.Cells(newrow, 12) = volume
                    newrow = newrow + 1
                    volume = 0
                    If ws.Cells(i + 1, 3) > 0 Then
                        opening = ws.Cells(i + 1, 3)
                    Else
                    End If
                End If
            Next i
        Next ws
End Sub

