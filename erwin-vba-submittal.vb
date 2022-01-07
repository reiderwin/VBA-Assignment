Sub StockData()

'Loop through all sheets'
Dim ws As Worksheet
For Each ws In Worksheets

ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Volume"

    'Determines LastRow as a variable
    Dim LastRow As Double

    'Determine total number of rows in current sheet
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Set initial variable for holding the ticker symbol
    Dim Ticker As String
    
    'Set initial varialbe for holding the total volume per Ticker
    Dim Total_Volume As Double
    Total_Volume = 0
    
    'Keep track of the location for each Ticker in the Summary Table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    'Set inital varialbe for Opening Price & Closing Price
    Dim Opening_Price As Double
    Dim Closing_Price As Double
    Opening_Price = ws.Cells(2, 3).Value
    
        'Loops through all Tickers
        For i = 2 To LastRow
        
            'Checks to see if within the same Ticker
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            'Sets the Ticker Symbol
            Ticker = ws.Cells(i, 1).Value
            
            'Prints Ticker Symbol to Summary Table
            ws.Range("I" & Summary_Table_Row).Value = Ticker
            
            'Adds to the Total Volume
            Total_Volume = Total_Volume + ws.Cells(i, 7).Value
            
            'Prints the Total_Volume to the Summary Table
            ws.Range("L" & Summary_Table_Row).Value = Total_Volume
            
            'Resets Total_Volume
            Total_Volume = 0
            
            'Sets the new Closing Price
            Closing_Price = ws.Cells(i, 6).Value
            
            'Sets Varialbe & Performs Calculation for Yearly Change
            Dim Yearly_Change As Double
            Yearly_Change = Closing_Price - Opening_Price
            
            'Prints Yearly Change to Summary Table
            ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
            
                'Performs loop to format color code for negative or positive
                If Yearly_Change < 0 Then
                
                'Changes cell color to Red for negative values
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                
                Else
                
                'Changes cell color to Green for zero or positive
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                
                End If
                
                
            'Sets Variable & Performs Calculation for Percent Change
            Dim Percent_Change As Double
            
                    'This loop catches when the Opening Price & Closing Price are equal to zero which causes an error
                    If Closing_Price And Opening_Price <> 0 Then
                    Percent_Change = (Closing_Price - Opening_Price) / Opening_Price
                    Else
                    Percent_Change = 0
                    End If
                    
            'Prints Percent Change to Summary Table
            ws.Range("K" & Summary_Table_Row).Value = Percent_Change
            
            'Formats Percent Change Column to be shown as Percentage
            ws.Cells(Summary_Table_Row, 11).Style = "Percent"
            
            'Add one to the Summary Table Row
            Summary_Table_Row = Summary_Table_Row + 1
            
            'Sets the new Opening Price
            Opening_Price = ws.Cells(i + 1, 3).Value
            
            'If the following cell is the same brand
            Else
            
            'Add to the Total_Volume
            Total_Volume = Total_Volume + ws.Cells(i, 7).Value
            
            End If
            
        Next i
        
Next ws

End Sub

