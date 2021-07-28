Sub Stocks_WorksheetLoop()

'ceate a script that will loop through all the stocks for one year and output

    Dim WS As Worksheet
    
    For Each WS In ActiveWorkbook.Worksheets
    
'create column headings
    WS.Range("I1").Value = "Ticker"
    WS.Range("J1").Value = "Yearly Change"
    WS.Range("K1").Value = "Percent Change"
    WS.Range("L1").Value = "Total Stock Volume"

'decalre all other varibles
    Dim Ticker As String
    Ticker = " "
    Dim Ticker_total As Double
    Ticker_total = 0
    
'declare other variables (the default value of Double (provides the largest and smallest possible magnitudes for a number) is 0)

    Dim Open_Price As Double
    Open_Price = 0
    Dim Closing_Price As Double
    Closing_Price = 0
    Dim Yearly_Change As Double
    Yearly_Change = 0
    Dim Percent_Change As Double
    Percent_Change = 0
    Dim Total_Stock_Volume As Double
    Total_Stock_Volume = 0
    Dim Table_Rows As Long
    Table_Rows = 2
    
'create row count for sheets

    'Dim Lastrow As Long
    LastRow = WS.Cells(Rows.Count, "A").End(xlUp).Row
    
    Dim i As Long
    For i = 2 To LastRow

'get ticker value
    Ticker = WS.Cells(i, 1).Value

'get the open price at the beginning of the year
    If Open_Price = 0 Then
        Open_Price = WS.Cells(i, 3).Value
    End If
    
'the total stock volume of the stock
'add values to total stock volume
    Total_Stock_Volume = Total_Stock_Volume + WS.Cells(i, 7).Value

'confirm ticker
    If WS.Cells(i + 1, 1).Value <> Ticker Then
    
'get increment for ticker totals for when tickers vary
    Ticker_total = Ticker_total + 1
    WS.Cells(Ticker_total + 1, 9) = Ticker

'get the closing price for the end of the year
    Closing_Price = WS.Cells(i, 6).Value

'get the yearly change
    Yearly_Change = Closing_Price - Open_Price

'print yearly change in column J
    WS.Cells(Ticker_total + 1, 10).Value = Yearly_Change
    
'get percent change value
    If Open_Price = 0 Then
        Percent_Change = 0
    Else
        Percent_Change = (Yearly_Change / Open_Price)
    End If

'format the percent change
    WS.Cells(Ticker_total + 1, 11).Value = Format(Percent_Change, "Percent")
      
'set open price to 0 if get a different ticker
    Open_Price = 0

'print the total stock volume in column L
    WS.Cells(Ticker_total + 1, 12).Value = Total_Stock_Volume
    
'set total stock volume to 0 if get a different ticker
    Total_Stock_Volume = 0
    

'conditional formatting that will highlight positive change in green and negative change in red
    If Yearly_Change > 0 Then
        WS.Cells(Ticker_total + 1, 10).Interior.Color = RGB(0, 255, 0)
    ElseIf Yearly_Change < 0 Then
        WS.Cells(Ticker_total + 1, 10).Interior.Color = RGB(255, 0, 0)
    End If
        
    End If

    Next i
    
    Next WS

End Sub
