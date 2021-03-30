Attribute VB_Name = "Module1"
Sub stock_calc()

'set initial variable for ticker symbol
Dim Ticker_Name As String

'set initial variables for yearly change and percent change
Dim Yearly_Change As Double
Dim Percent_Change As Double

'*****************************
'BONUS
'set greatest % increase/decrease/volume variables
'*****************************
Dim Greatest_Percent_Increase As Double
Dim Greatest_Percent_Decrease As Double
Dim Greatest_Total_Volume As Double
Dim Greatest_Percent_Increase_Ticker As String
Dim Greatest_Percent_Decrease_Ticker As String
Dim Greatest_Total_Volume_Ticker As String
Greatest_Percent_Increase = 0
Greatest_Percent_Decrease = 0
Greatest_Total_Volume = 0

'******************************
'BONUS
'Loop through all sheets
'******************************

For Each ws In Worksheets

    'Create variable to hold file name
    'Dim WorksheetName As String

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
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

    ' BONUS: counts the number of rows
    Dim lastrow As Long
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    '***************************************
    'autofit the columns widths of active sheet
    'ActiveSheet.Columns("I:L").AutoFit
     ws.Columns("I:L").AutoFit
    '***************************************
    
    'FORMAT % cells
    ws.Range("K2:K" & lastrow).NumberFormat = "0.00%"

    'Loop through all stock data
    For i = 2 To lastrow
       
       'Check if still within the same ticker symbol, if it is not...
       If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
         'Set the ticker symbol name to next ticker name and print to table
         Ticker_Name = ws.Cells(i, 1).Value
         ws.Range("I" & Summary_Table_Row) = Ticker_Name
         
         
         'Set the final value of Total Stock Volume and print to table
          Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
          ws.Range("L" & Summary_Table_Row) = Total_Stock_Volume
         
         'Calulate Yearly_Change and print to table
         Year_Close_Price = ws.Cells(i, 6).Value
         Yearly_Change = Year_Close_Price - Year_Open_Price
         ws.Range("J" & Summary_Table_Row) = Yearly_Change
         
         'highlight positive(green) and negative(red) yearly changes
         If Yearly_Change < 0 Then
            ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3
         End If
         
         If Yearly_Change > 0 Then
            ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4
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
            
         ws.Range("K" & Summary_Table_Row) = Percent_Change
         
         'Reset Total Stock Volume
         Total_Stock_Volume = 0
         
         'Set year open price for next ticker name
         Year_Open_Price = ws.Cells(i + 1, 3)
         
         'Add one to the summary table row
         Summary_Table_Row = Summary_Table_Row + 1
         
       Else
       
         'Add to the Total Stock Volume
         Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
         
      End If
      
       'Greatest variables set
       If Percent_Change > Greatest_Percent_Increase Then
          Greatest_Percent_Increase = Percent_Change
          Greatest_Percent_Increase_Ticker = Ticker_Name
       End If
       
       If Percent_Change < Greatest_Percent_Decrease Then
          Greatest_Percent_Decrease = Percent_Change
          Greatest_Percent_Decrease_Ticker = Ticker_Name
       End If
       
       If Total_Stock_Volume > Greatest_Total_Volume Then
          Greatest_Total_Volume = Total_Stock_Volume
          Greatest_Total_Volume_Ticker = Ticker_Name
      End If
      
     Next i
    
     
  Next ws
  
    '***************************************
    'autofit the columns widths of active sheet
    'ActiveSheet.Columns("I:L").AutoFit
     Columns("O:Q").AutoFit
    '***************************************
    
    'FORMAT % cells
    Range("Q2:Q3").NumberFormat = "0.00%"
     
     'Create headers for summary tables
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    
    Range("O2").Value = "Greatest % Increase"
    Range("P2").Value = Greatest_Percent_Increase_Ticker
    Range("Q2").Value = Greatest_Percent_Increase
    
    Range("O3").Value = "Greatest % Decrease"
    Range("P3").Value = Greatest_Percent_Decrease_Ticker
    Range("Q3").Value = Greatest_Percent_Decrease
    
    Range("O4").Value = "Greatest Total Volume"
    Range("P4").Value = Greatest_Total_Volume_Ticker
    Range("Q4").Value = Greatest_Total_Volume
    
    
     
End Sub

    

