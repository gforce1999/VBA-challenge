Attribute VB_Name = "Module1"
Sub stock_calc()

'set initial variable for ticker symbol
Dim Ticker_Name As String

'set initial variables for yearly change and percent change
Dim Yearly_Change As Double
Dim Percent_Change As Double

'set initial variables for year open price and year closing price
Dim Year_Open_Price As Double
Year_Open_Price = Cells(2, 3)

Dim Year_Close_Price As Double


'set inital variable for holding total stock volume per ticker
Dim Total_Stock_Volume As Double
Total_Stock_Volume = 0

'Keep track of the location for each ticker symbol name in the summary table
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

'Create headers for summary tables
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"

' BONUS: counts the number of rows
Dim lastrow As Long
lastrow = Cells(Rows.Count, 1).End(xlUp).Row

'***************************************
'autofit the columns widths of active sheet
ActiveSheet.Columns("I:L").AutoFit
'***************************************

'FORMAT % cells
Range("K2:K" & lastrow).NumberFormat = "0.00%"

'Loop through all stock data
For i = 2 To lastrow
   
   'Check if still within the same ticker symbol, if it is not...
   If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
     'Set the ticker symbol name to next ticker name and print to table
     Ticker_Name = Cells(i, 1).Value
     Range("I" & Summary_Table_Row) = Ticker_Name
     
     
     'Set the final value of Total Stock Volume and print to table
      Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
      Range("L" & Summary_Table_Row) = Total_Stock_Volume
     
     'Calulate Yearly_Change and print to table
     Year_Close_Price = Cells(i, 6).Value
     Yearly_Change = Year_Close_Price - Year_Open_Price
     Range("J" & Summary_Table_Row) = Yearly_Change
     
     'highlight positive(green) and negative(red) yearly changes
     If Yearly_Change < 0 Then
        Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3
     End If
     
     If Yearly_Change > 0 Then
        Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4
     End If
     
     
     'Calculate Percent Change and print to table
     If Year_Open_Price <> 0 Then
        Percent_Change = ((Year_Close_Price - Year_Open_Price) / Year_Open_Price)
     Else
        Percent_Change = Year_Close_Price - Year_Open_Price
     End If
        '**************************************
        'error handling specifically for PLNT which is 0 in 2014
        '**************************************
        
     Range("K" & Summary_Table_Row) = Percent_Change
     
     'Reset Total Stock Volume
     Total_Stock_Volume = 0
     
     'Set year open price for next ticker name
     Year_Open_Price = Cells(i + 1, 3)
     
     'Add one to the summary table row
     Summary_Table_Row = Summary_Table_Row + 1
     
   Else
   
     'Add to the Total Stock Volume
     Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
     
  End If
  
 Next i
 
End Sub

    

