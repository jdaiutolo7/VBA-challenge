VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub StockMarketData()
    
    For Each ws In Worksheets

        ' Establish the column headers and place them in the appropriate cells
        ws.Range("I1") = "Ticker"
        ws.Range("J1") = "Yearly Change"
        ws.Range("K1") = "Percent Change"
        ws.Range("L1") = "Total Stock Volume"
    
        ' Define the column headers as Variables
        Dim Ticker_Symbol As String
        Dim Yearly_Change As Double
        Dim Percent_Change As Double
        Dim Total_Stock_Volume As Double
        
        ' Define Opening and Closing Prices
        Dim Open_Price As Double
        Dim Close_Price As Double
    
    ' Setting the Total Stock Volume so it's initially 0
    Total_Stock_Volume = 0
    
    Dim Summary_Table_Row As Double
    Summary_Table_Row = 2
    
    ' Identifying the last row function as "lastrow" so it's easier to type going forward
    Dim lastrow As Double
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
    ws.Columns("A:Q").AutoFit
    
    For i = 2 To lastrow
        
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1) Then
            Ticker_Symbol = ws.Cells(i, 1).Value
            Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
            
            ws.Range("I" & Summary_Table_Row).Value = Ticker_Symbol
            ws.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
            
            Total_Stock_Volume = 0
            
            Close_Price = ws.Cells(i, 6)
            
        If Open_Price = 0 Then
            Yearly_Change = 0
            Percent_Change = 0
        
        Else
            Yearly_Change = Close_Price - Open_Price
            Percent_Change = (Close_Price - Open_Price) / Open_Price
    
    End If
    
        ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
        ws.Range("K" & Summary_Table_Row).Value = Percent_Change
        ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
        Summary_Table_Row = Summary_Table_Row + 1
    
        ElseIf ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1) Then
            Open_Price = ws.Cells(i, 3)
        
        Else
            Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
    
      End If
    
   Next i

 For j = 2 To lastrow

    If ws.Range("J" & j).Value > 0 Then
       ws.Range("J" & j).Interior.ColorIndex = 4

    ElseIf ws.Range("J" & j).Value < 0 Then
        ws.Range("J" & j).Interior.ColorIndex = 3
        
    End If

 Next j
    
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"

    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"

    Dim Greatest_Increase As Double
    Dim Greatest_Decrease As Double
    Dim Greatest_Volume As Double

    Greatest_Increase = 0
    Greatest_Decrease = 0
    Greatest_Volume = 0

For k = 2 To lastrow

    If ws.Cells(k, 11).Value > Greatest_Increase Then
        Greatest_Increase = ws.Cells(k, 11).Value
        ws.Range("Q2").Value = Greatest_Increase
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("P2").Value = ws.Cells(k, 9).Value
        
    End If

  Next k

For n = 2 To lastrow
    
    If ws.Cells(n, 11).Value < Greatest_Decrease Then
        Greatest_Decrease = ws.Cells(n, 11).Value
        ws.Range("Q3").Value = Greatest_Decrease
        ws.Range("Q3").NumberFormat = "0.00%"
        ws.Range("P3").Value = ws.Cells(n, 9).Value
        
    End If
    
 Next n

For m = 2 To lastrow
    
    If ws.Cells(m, 12).Value > Greatest_Volume Then
        Greatest_Volume = ws.Cells(m, 12).Value
        ws.Range("Q4").Value = Greatest_Volume
        ws.Range("P4").Value = ws.Cells(m, 9).Value
        
    End If
  
  Next m
    
Next ws

End Sub


