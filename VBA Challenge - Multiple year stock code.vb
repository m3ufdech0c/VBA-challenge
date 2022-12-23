Option Explicit

Public Sub SummarizeStock()

Dim CurrSheet, StkTicker
Application.ScreenUpdating = False


' Loop through all the sheets in the workbook
For Each CurrSheet In Worksheets
    CurrSheet.Activate
    
    '----------------------------------------------------------------------------
    '------ Start the calcualtion for summary table -----------------------------
    '----------------------------------------------------------------------------
    
    ' Declare variable and hold the last row number for the current sheet
    Dim last As Long
    last = CurrSheet.Cells(Rows.Count, "G").End(xlUp).Row
    
    ' Iterator variable
    Dim i, j As Long
    
    ' Declare variable to hold ticker name
    Dim Ticker_Name As String
    
    ' Declare and set variable to hold open price
    ' The first value is the open price of the first stock
    Dim Open_Price As Double
    Open_Price = Cells(2, 3)
    
    ' Declare and initialize variable for holding stock volume total
    Dim Volume_Stock As Double
    Volume_Stock = 0
    
    ' Declare and set variable to hold close price
    Dim Close_Price As Double
    Close_Price = 0
    
    ' Keep track of the location for each stock ticker
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    ' Set summary table headers
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
    ' Loop through all stocks records
    For i = 2 To last
        
        ' Check if we are still within the same stock ticker, if it is not...
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                       
            Ticker_Name = Cells(i, 1).Value
            
            ' Increase the volume of stock
            Volume_Stock = Volume_Stock + Cells(i, 7).Value
            
            Close_Price = Cells(i, 6).Value
            
            ' Assign Ticker Name
            Range("I" & Summary_Table_Row).Value = Ticker_Name
            
            With Range("J" & Summary_Table_Row)
                ' Assign Yearly change
                .Value = Close_Price - Open_Price
                ' Delete Existing Conditional Formatting from Range
                .FormatConditions.Delete
                ' Add formatting Red for negative values
                .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
                    Formula1:="=0"
                .FormatConditions(1).Interior.Color = RGB(255, 0, 0)
                
                .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
                    Formula1:="=0"
                .FormatConditions(2).Interior.Color = RGB(0, 255, 0)
            End With
            
            ' Calculate percentage change and assign to cell
            Range("K" & Summary_Table_Row).Value = (Close_Price - Open_Price) / Open_Price
            Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
            
            ' Assign volume stock total
            Range("L" & Summary_Table_Row).Value = Volume_Stock
            
            
            
          ' Add one to the summary table row
          Summary_Table_Row = Summary_Table_Row + 1
          
          ' Assign the Opening Price for next Ticker
          Open_Price = Cells(i + 1, 3).Value
          
          ' Reset Volume stock accumulator for next stock
          Volume_Stock = 0
    
        
        Else
            Volume_Stock = Volume_Stock + Cells(i, 7).Value
        End If
    
    Next i
    
    '----------------------------------------------------------------------------
    '------ Start the calcualtion for outlier -----------------------------------
    '----------------------------------------------------------------------------

    ' Declare variable and hold the last row number for the current sheet summary table
    Dim last1 As Long
    last1 = CurrSheet.Cells(Rows.Count, "I").End(xlUp).Row
    
    'Declare and initalize variables to hold greatest percent increase
    Dim Greatest_Increase_Ticker As String
    Greatest_Increase_Ticker = Cells(2, 9).Value
    
    Dim Greatest_Increase_Percent As Double
    Greatest_Increase_Percent = Cells(2, 11).Value
    
    'Declare and initalize variables to hold greatest percent decrease
    Dim Greatest_Decrease_Ticker As String
    Greatest_Decrease_Ticker = Cells(2, 9).Value
    
    Dim Greatest_Decrease_Percent As Double
    Greatest_Decrease_Percent = Cells(2, 11).Value
    
    'Declare and initalize variables to hold greatest stock volume
    Dim Greatest_TotVol_Ticker As String
    Greatest_TotVol_Ticker = Cells(2, 9).Value
    
    Dim Greatest_TotVol_Value As Double
    Greatest_TotVol_Value = Cells(2, 12).Value
    
    ' Loop through summary table
    For j = 2 To last1 - 1
        
        ' Test if next row percent increase is greater than current swap to variable if true
        If Cells(j + 1, 11).Value > Greatest_Increase_Percent Then
            Greatest_Increase_Ticker = Cells(j + 1, 9).Value
            Greatest_Increase_Percent = Cells(j + 1, 11).Value
        
        End If
        
        ' Test if next row percent increase is lower than current swap to variable if true
        If Cells(j + 1, 11).Value < Greatest_Decrease_Percent Then
            Greatest_Decrease_Ticker = Cells(j + 1, 9).Value
            Greatest_Decrease_Percent = Cells(j + 1, 11).Value
        End If
        
        ' Test if next row total volume is greater than current swap to variable if true
        If Cells(j + 1, 12).Value > Greatest_TotVol_Value Then
            Greatest_TotVol_Ticker = Cells(j + 1, 9).Value
            Greatest_TotVol_Value = Cells(j + 1, 12).Value
        End If
        
    Next j
    
    
    ' Assign variables to cells
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    
    Range("O2").Value = "Greatest % Increase"
    Range("P2").Value = Greatest_Increase_Ticker
    Range("Q2").Value = Greatest_Increase_Percent
    Range("Q2").NumberFormat = "0.00%"
    
    Range("O3").Value = "Greatest % Decrease"
    Range("P3").Value = Greatest_Decrease_Ticker
    Range("Q3").Value = Greatest_Decrease_Percent
    Range("Q3").NumberFormat = "0.00%"
    
    Range("O4").Value = "Greatest Total Volume"
    Range("P4").Value = Greatest_TotVol_Ticker
    Range("Q4").Value = Greatest_TotVol_Value
    
    
Next

Application.ScreenUpdating = True
End Sub
