Attribute VB_Name = "Module1"
Sub stockrpt()
     
    ' Loop through worksheets
    Dim ws As Worksheet
    
    For Each ws In Worksheets
        ws.Activate
      
     'define variables
    Dim s As Long
    Dim r As Long
    Dim c As Long
    Dim i As Long
    Dim stock_vol As Double
    Dim yr_chg As Double
    Dim per_chg As Double
    Dim opening As Double
    Dim closing As Double
    Dim greatest_vol As Double
    Dim greatest_inc As Double
    Dim greatest_dec As Double
    Dim lastrow As Long
    
    ' initialize fields
    s = 2
    curr_ticker = "  "
    greatest_inc = 0
    greatest_dec = 0
    greatest_vol = 0

  
         
    'count the number of rows
     lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
    'Create headers
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Stock Volume"
    Cells(2, 14).Value = "Greatest % Increase"
    Cells(3, 14).Value = "Greatest % Decrease"
    Cells(4, 14).Value = "Greatest total Volume"
    Cells(1, 15).Value = "Ticker"
    Cells(1, 16).Value = "Value"
    'Check for ticker
    For r = 2 To lastrow
        If curr_ticker = "  " Then
            'get opening from 1st record for ticker
            opening = Cells(r, 3).Value
            curr_ticker = Cells(r, 1).Value
        End If
        'calculate the total stock value sum of all the volumes
        stock_vol = stock_vol + Cells(r, 7).Value
        'last row of current stock
        If Cells(r + 1, 1).Value <> curr_ticker Then
            'get closing for last record for ticker
            closing = Cells(r, 6).Value
            'compute difference & write it red for negative green for positive
            yr_chg = closing - opening
            Cells(s, 10).Value = yr_chg
            If Cells(s, 10).Value < 0 Then
                Cells(s, 10).Interior.ColorIndex = 3
            ElseIf Cells(s, 10).Value > 0 Then
                Cells(s, 10).Interior.ColorIndex = 4
            End If
            'compute the percent change
            per_chg = yr_chg / opening
            ' move the values to the fields
            Cells(s, 9).Value = curr_ticker
            
            Cells(s, 11).Value = FormatPercent(per_chg)
            Cells(s, 12).Value = stock_vol
          
            ' set fields for next ticker
            s = s + 1
            opening = Cells(r + 1, 3).Value
            curr_ticker = Cells(r + 1, 1).Value
            stock_vol = 0
        End If
    Next r

    For i = 2 To lastrow
        If Cells(i, 11).Value > greatest_inc Then
            greatest_inc = Cells(i, 11).Value
            inc_ticker = Cells(i, 9).Value
        End If
        If Cells(i, 11).Value < greatest_dec Then
            greatest_dec = Cells(i, 11).Value
            dec_ticker = Cells(i, 9).Value
        End If
        If Cells(i, 12).Value > greatest_vol Then
            greatest_vol = Cells(i, 12).Value
            vol_ticker = Cells(i, 9).Value
        End If
     Next i
    Cells(2, 16).Value = FormatPercent(greatest_inc)
    Cells(2, 15).Value = inc_ticker
    Cells(3, 16).Value = FormatPercent(greatest_dec)
    Cells(3, 15).Value = dec_ticker
    Cells(4, 16).Value = greatest_vol
    Cells(4, 15).Value = vol_ticker
    
 
    Next ws
End Sub


