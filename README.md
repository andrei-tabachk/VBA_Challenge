# VBA_Challenge




' run on full workbook
Sub wsfullrun()
    
    Dim ws As Worksheet
    
    Application.ScreenUpdating = False
    For Each ws In Worksheets
        ws.Select
        Call TickerChallenge
    Next
    
    Application.ScreenUpdating = True
    
End Sub



Sub TickerChallenge():

    'Insert Data Via Range and set as column headers
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"


    'Variables for tickers , sum , percent change , yearly change
    Dim ticker As String
    Dim next_ticker As String
    Dim volume_total As Double
    Dim percent As Double
    Dim year_chg As Double
    Dim openprice As Double
    Dim closeprice As Double
    Dim opendate As Double
    Dim closedate As Double
    Dim lastrow As Long
    Dim prev_ticker As String
    Dim maxinc As Double
    Dim maxdec As Double
    Dim maxvol As Double
    
    maxinc = 0
    maxdec = 0
    maxvol = 0
    
    'opendate
    opendate = 20200102
    closedate = 20201231
    
    'Total per ticker
    volume_total = 0
    
    ' Keep track of the location for each credit card brand in the summary table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    'lastrow finder
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
    'loop for tickers
    For i = 2 To lastrow

    ticker = Cells(i, 1).Value
    next_ticker = Cells(i + 1, 1).Value
    prev_ticker = Cells(i - 1, 1).Value

    If next_ticker <> ticker Then
    
      ' Print the ticker
      Cells(Summary_Table_Row, 9).Value = ticker
      
      'print total amount
      Cells(Summary_Table_Row, 12).Value = volume_total

      ' Add to the ticker total
      
      volume_total = volume_total + Cells(i, 7).Value
      
      'closing price
      closeprice = Cells(i, 6).Value
      
          
      ' add yearly change, beginning of the year vs close at end of year
      year_chg = closeprice - openprice
      
      'print price change
      Cells(Summary_Table_Row, 10).Value = year_chg
      
      'add percent change between opening price of beginning of the year vs at close end of year
      percent = (year_chg / openprice)
      Cells(Summary_Table_Row, 11).NumberFormat = "0.00%"
      
      'print percent
      Cells(Summary_Table_Row, 11).Value = percent
      
      'add another row
       Summary_Table_Row = Summary_Table_Row + 1
    
      ' If the cell immediately following a row is the same ticker...

    
    ElseIf ticker <> prev_ticker Then
    
    'opening price
    openprice = Cells(i, 3).Value
    
    volume_total = Cells(i, 7).Value
    
    Else
      ' Add to the ticker Total
      volume_total = volume_total + Cells(i, 7).Value
      
    End If

  Next i


    For k = 2 To lastrow
    
        If Cells(k, 11).Value > maxinc Then
            maxinc = Cells(k, 11).Value
            ticker = Cells(k, 9).Value
            Cells(2, 16).Value = ticker
            Cells(2, 17).Value = maxinc
            Cells(2, 17).NumberFormat = "0.00%"
        ElseIf Cells(k, 11).Value < maxdec Then
            maxdec = Cells(k, 11).Value
            ticker = Cells(k, 9).Value
            Cells(3, 16).Value = ticker
            Cells(3, 17).Value = maxdec
            Cells(3, 17).NumberFormat = "0.00%"
        ElseIf Cells(k, 12).Value > maxvol Then
            maxvol = Cells(k, 12).Value
            ticker = Cells(k, 9).Value
            Cells(4, 16).Value = ticker
            Cells(4, 17).Value = maxvol
    End If
    
    Next k


Call color

End Sub

Sub color()



 ' Keep track of the location for each credit card brand in the summary table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    'lastrow finder
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
    'loop for tickers
    For i = 2 To lastrow
    year_chg = Cells(i, 10).Value
    
        If year_chg >= 0 Then
            Cells(i, 10).Interior.ColorIndex = 4
        ElseIf year_chg < 0 Then
            Cells(i, 10).Interior.ColorIndex = 3
        End If
    
    Next i
    
End Sub
