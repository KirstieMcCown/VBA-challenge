Sub VBAWallStreet():

'Set a initial variable for the ticker symbol
Dim Ticker_Symbol As String

'Set a initial variable for the total stock volume
Dim Stock_Volume As Double
Stock_Volume = 0

'Set a variable for the opening price at the beginning of the year
Dim Opening_Price As Double

'Set a variable for the closing price at the end of the year
Dim Closing_Price As Double

'Set a variable for the Yearly Change
Dim Yearly_Change As Double

'Set a variable for the Percent Change - Check this
Dim Percent_Change As Double

'Define new location for each ticker symbol
Dim Combined_Ticker As Integer
Combined_Ticker = 2

'Define Last Row and Last Column
LastRow = Cells(Rows.Count, 1).End(xlUp).Row
LastColumn = Cells(1, Columns.Count).End(xlToLeft).Column

'Loop through all Ticker Symbols

For i = 2 To LastRow

    'Check down the column, to see if each cell has any of the same ticker symbols
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

    'Set the Ticker Symbol Variable
    Ticker_Symbol = Cells(i, 1).Value

     'Add the Total of Stock Volume for each ticker symbol
    Stock_Volume = Stock_Volume + Cells(i, 7).Value

    'Print the Ticker Symbol in the Summary Table
    Range("I" & Combined_Ticker).Value = Ticker_Symbol

    'Print the Total Stock Volume in the Summary Table
    Range("L" & Combined_Ticker).Value = Stock_Volume

    'Add one to the Combined Ticker Location
    Combined_Ticker = Combined_Ticker + 1
    'Reset the Total Stock Volume

    Stock_Volume = 0
    'If the cell immediately following the last row is the same Ticker Symbol

    Else

    'Add to the Total Stock Volume
    Stock_Volume = Stock_Volume + Cells(i, 7).Value

  End If

        'Check down the column, to see if each cell has any of the same ticker symbols
        If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then

        'Set the Opening Price Variable
        Opening_Price = Cells(i, 3).Value

        'Print opening price to check correct value is being stored
        'Range("J" & Combined_Ticker).Value = Opening_Price

        'Check down the column, to see if each cell has any of the same ticker symbols
        ElseIf Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

        'Set the Closing Price Variable
        Closing_Price = Cells(i, 6).Value

        'Print closing price to check correct value is being stored
        'Range("K" & Combined_Ticker - 1).Value = Closing_Price

        'Print the Yearly Change in the Summary Table
        Range("J" & Combined_Ticker - 1).Value = Closing_Price - Opening_Price

        'Define the Yearly Change Variable
        'Yearly_Change = Cells(i, 10).Value

        If Opening_Price = 0 Then
        Range("K" & Combined_Ticker - 1).Value = "N/A"

        Else

        'Calculate and Print the Percent Change in the Summary Table - change value to percentage via VBA
        Range("K" & Combined_Ticker - 1).NumberFormat = "00.00%"
        Range("K" & Combined_Ticker - 1).Value = ((Closing_Price - Opening_Price) / Opening_Price)

           End If
           End If

        'Check if the Percent Change is an increase
        If Range("K" & Combined_Ticker - 1).Value > 0 Then

        'Colour the positive percentage increase green
          Range("K" & Combined_Ticker - 1).Interior.ColorIndex = 4

          'Check if the Percent Change is an decrease
        ElseIf Range("K" & Combined_Ticker - 1).Value < 0 Then

        'Colour the negative percentage decrease red
        Range("K" & Combined_Ticker - 1).Interior.ColorIndex = 3

    End If

Next i

'Print column headers in each sheet
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        
End Sub