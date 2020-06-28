Attribute VB_Name = "Module1"
Sub Stocks():

    Dim Ticker_Symbol As String
    Dim Opening_Price As Double
    Dim Closing_Price As Double
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    Dim Total_Stock_Volume As Double
    Dim Greatest_Increase As Double
    Dim Greatest_Increase_Ticker As String
    Dim Greatest_Decrease As Double
    Dim Greatest_Decrease_Ticker As String
    Dim Greatest_Total_Volume As Double
    Dim Greatest_Total_Volume_Ticker As String
    Dim Bad_Data As Boolean
    
    Dim LastRow As Long
    Dim SummaryRow As Long
    
    For Each ws In Worksheets
        ' Labels
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        ' Initialize some values
        SummaryRow = 2
        Total_Stock_Volume = 0
        Opening_Price = ws.Cells(2, 3).Value
        Greatest_Increase = 0
        Greatest_Decrease = 0
        Greatest_Total_Volume = 0
        BadData = False
    
        ' Determine the Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row + 1
            
        For i = 2 To LastRow
        Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7)
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                ' Ticker
                Ticker_Symbol = ws.Cells(i, 1).Value
                ws.Range("I" & SummaryRow).Value = Ticker_Symbol
                
                ' Yearly Change
                If Not Bad_Data Then
                    Closing_Price = ws.Cells(i, 6).Value
                    Yearly_Change = Closing_Price - Opening_Price
                    ws.Range("J" & SummaryRow).Value = Yearly_Change
                    If Yearly_Change > 0 Then
                        ws.Range("J" & SummaryRow).Interior.ColorIndex = 4
                    Else
                        ws.Range("J" & SummaryRow).Interior.ColorIndex = 3
                    End If
                Else
                    ' potentially put something in cell to denote bad data
                End If
                For j = i + 1 To LastRow
                    ' Avoiding potential divide by zero by looking for nonzero opening price
                    
                    If ws.Cells(j, 3).Value <> 0 Then
                        Opening_Price = ws.Cells(j, 3).Value
                        Bad_Data = False
                        Exit For
                    End If
                    If ws.Cells(j + 1, 1).Value <> ws.Cells(j, 1).Value Then
                        Bad_Data = True
                        Exit For
                    End If
                Next j
                
                'Percent Change
                If Not Bad_Data Then
                    Percent_Change = Yearly_Change / Opening_Price
                    ws.Range("K" & SummaryRow).Value = Percent_Change
                    ws.Range("K" & SummaryRow).NumberFormat = "0.00%"
                Else
                    ' potentially put something in cell to denote bad data
                End If
                
                ' Total Stock Volume
                ws.Range("l" & SummaryRow).Value = Total_Stock_Volume
                
                ' Greatest Values Evaluations
                If Greatest_Increase < Percent_Change Then
                    Greatest_Increase_Ticker = Ticker_Symbol
                    Greatest_Increase = Percent_Change
                End If
                If Percent_Change < Greatest_Decrease Then
                    Greatest_Decrease_Ticker = Ticker_Symbol
                    Greatest_Decrease = Percent_Change
                End If
                If Greatest_Total_Volume < Total_Stock_Volume Then
                    Greatest_Total_Volume_Ticker = Ticker_Symbol
                    Greatest_Total_Volume = Total_Stock_Volume
                End If
                
                Total_Stock_Volume = 0
                SummaryRow = SummaryRow + 1
            End If
            
        Next i
        
        ' Greatest Values Reported
        ws.Range("P2").Value = Greatest_Increase_Ticker
        ws.Range("Q2").Value = Greatest_Increase
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("P3").Value = Greatest_Decrease_Ticker
        ws.Range("Q3").Value = Greatest_Decrease
        ws.Range("Q3").NumberFormat = "0.00%"
        ws.Range("P4").Value = Greatest_Total_Volume_Ticker
        ws.Range("Q4").Value = Greatest_Total_Volume
        
        ' Autofit to display data
        ws.Columns("I:Q").AutoFit
    Next ws

End Sub

