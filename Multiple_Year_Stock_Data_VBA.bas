Attribute VB_Name = "Module1"
' VBA_Multiple_Year_Stock_Data

Sub Stock_Data()

    'Loop Through All Sheets
    For Each ws In Worksheets
    
        'Declare the data type of a variable
        Dim Ticker_Symbol As String
        Dim Yearly_Change As Double
        Dim Percent_Change As Double
        Dim Stock_Volume As Double
        Dim Open_Price As Double
        Dim Close_Price As Double
        Dim Summary_Table_Row As Integer
        Dim LastRow As Double
        
        'Set an initial variable
        Ticker_Symbol = 0
        Yearly_Change = 0
        Percent_Change = 0
        Stock_Volume = 0
        Open_Price = 0
        Close_Price = 0
        Summary_Table_Row = 2
        
        'Set Summary Table Header
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "percent change"
        ws.Cells(1, 12).Value = "stock volume"
        
        'Autofit to display data
        ws.Columns("A:Q").AutoFit
        
        'Determine the Last Row For Each Sheet
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        

        
        'Loop Through All Stock
        For i = 2 To LastRow
            
            'Open Price Each Year For Each Stock
            If Open_Price = 0 Then
            Open_Price = ws.Cells(i, 3).Value
            End If
            
            'Output The Total Stock Volume Of The Stock
            Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value

            'Output The Ticker Symbol
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                Ticker_Symbol = ws.Cells(i, 1).Value
                
                'Close Price Each Year For Each Stock
                Close_Price = ws.Cells(i, 6)
            
                'Calculate Yearly change
                Yearly_Change = Close_Price - Open_Price

                'Calculate The percent change
                If Open_Price = 0 Then
                Percent_Change = 0
                Else
                Percent_Change = Yearly_Change / Open_Price
                End If
                
                'Print The Values On The Summary Table
                ws.Range("j" & Summary_Table_Row).Value = Yearly_Change
                ws.Range("I" & Summary_Table_Row).Value = Ticker_Symbol
                ws.Range("K" & Summary_Table_Row).Value = Format(Percent_Change, "percent")
                ws.Range("L" & Summary_Table_Row).Value = Stock_Volume
                
                
                ' Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
                
                ' Reset the Values
                Open_Price = 0
                Stock_Volume = 0
          
          
            End If

    
        Next i
        
        
        'conditional formatting for Yearly Change
        Dim LastRow2 As Double
        LastRow2 = ws.Cells(Rows.Count, 10).End(xlUp).Row
    
        For j = 2 To LastRow2
                If ws.Cells(j, 10).Value > 0 Then
                    ws.Cells(j, 10).Interior.ColorIndex = 4
                ElseIf ws.Cells(j, 10).Value < 0 Then
                    ws.Cells(j, 10).Interior.ColorIndex = 3
                Else
                    ws.Cells(j, 10).Interior.ColorIndex = 2
                End If
        Next j
        
        'Bonus
        
        'Set The Greatest Table Header
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % increase"
        ws.Cells(3, 15).Value = "Greatest % decrease"
        ws.Cells(4, 15).Value = "Greatest total volume"
        
        
        'Print The Values On The Greatest Table
        Greatest_Inc_Value = ws.Cells(2, 17).Value
        Greatest_Dec_Value = ws.Cells(3, 17).Value
        Greatest_Tot_Vol_Value = ws.Cells(4, 17).Value
        Greatest_Inc_Ticker = ws.Cells(2, 16).Value
        Greatest_Dec_Ticker = ws.Cells(3, 16).Value
        Greatest_Tot_Vol_Ticker = ws.Cells(4, 16).Value
        
        
        'Determine the Last Row For Ticker Column for Each Sheet
        Dim LastRow3 As Double
        LastRow3 = ws.Cells(Rows.Count, 10).End(xlUp).Row
        
        'Loop Through All Ticker
        For k = 2 To LastRow3
        
            'Find The Greatest % increase
            If ws.Cells(k, 11) > Greatest_Inc_Value Then
                Greatest_Inc_Value = ws.Cells(k, 11).Value
                Greatest_Inc_Ticker = ws.Cells(k, 9).Value
            End If
            
            'Find The Greatest % decrease
            If ws.Cells(k, 11) < Greatest_Dec_Value Then
                Greatest_Dec_Value = ws.Cells(k, 11).Value
                Greatest_Dec_Ticker = ws.Cells(k, 9).Value
            End If
            
            'Find Greatest total volume
            If ws.Cells(k, 12) > Greatest_Tot_Vol_Value Then
                Greatest_Tot_Vol_Value = ws.Cells(k, 12).Value
                Greatest_Tot_Vol_Ticker = ws.Cells(k, 9).Value
            End If
            
        Next

            
            
            'conditional formatting for Greatest %
            ws.Cells(2, 17).Value = Format(Greatest_Inc_Value, "percent")
            ws.Cells(3, 17).Value = Format(Greatest_Dec_Value, "percent")
            ws.Cells(4, 17).Value = Greatest_Tot_Vol_Value
            
            ws.Cells(2, 16).Value = Greatest_Inc_Ticker
            ws.Cells(3, 16).Value = Greatest_Dec_Ticker
            ws.Cells(4, 16).Value = Greatest_Tot_Vol_Ticker

        
        
            
    Next ws
    
 
    
End Sub







