Attribute VB_Name = "Module2"
Sub StockMarket5()
' Create a script that loops through all the stocks for one year and outputs the following information:
'   * The ticker symbol.
'   * Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
'   * The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
'   * The total stock volume of the stock.
'
' **Note:** Make sure to use conditional formatting that will highlight positive change in green and negative change in red.
'
' ## Bonus
'   * Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume".
'   * Make the appropriate adjustments to your VBA script to allow it to run on every worksheet (that is, every year) just by running the VBA script once.
'
' ## Other Considerations
'   * Use the sheet `alphabetical_testing.xlsx` while developing your code. This data set is smaller and will allow you to test faster.
'   * Your code should run on this file in less than 3 to 5 minutes.
'   * Make sure that the script acts the same on every sheet. The joy of VBA is that it takes the tediousness out of repetitive tasks with one click of a button
    
    Application.ScreenUpdating = False ' Prevents visual updates while code is running
        
    ' Declare Ticker, Open, Close and Volume Variables
    Dim Ticker As String, TickerOpen As Double, TickerClose As Double, TickerVolume As Double
    
    ' Declare Greatest Ticker Names and Value Variables
    Dim TickerGreatestIncrease As String, TickerGreatestIncreaseValue As Double
    Dim TickerGreatestDecrease As String, TickerGreatestDecreaseValue As Double
    Dim TickerGreatestVolume As String, TickerGreatestVolumeValue As Double
    
    ' Declare variable to for displaying info
    Dim rowDisplay As Integer
    
    'Dim LastRow As Double
        
    ' Loop through each sheet
    For Each ws In ActiveWorkbook.Sheets
        ws.Activate ' Activate current worksheet...preventing using 'ws.' for every cell reference
                        
        ' Calculate last row of data
        LastRow = Range("A" & Rows.Count).End(xlUp).Row
        
        ' Display Column Titles
        Cells(1, 9) = "Ticker": Cells(1, 10) = "Yearly Change": Cells(1, 11) = "Percent Change": Cells(1, 12) = "Total Stock Volume"
        
        ' Initialize starting row for display information
        rowDisplay = 2
    
        ' Reset variables to hold Greatest Information
        TickerGreatestIncrease = "":        TickerGreatestIncreaseValue = -10000
        TickerGreatestDecrease = "":        TickerGreatestDecreaseValue = 10000
        TickerGreatestVolume = "":          TickerGreatestVolumeValue = -1
        
        ' Cycle through each row
        For rowcounter = 2 To LastRow
            If rowcounter = 2 Then
                Ticker = Cells(rowcounter, 1)
                TickerOpen = Cells(rowcounter, 3)
            End If
            
            ' Determine if row is last row for specific Ticker
            If Cells(rowcounter + 1, 1) <> Cells(rowcounter, 1) Then
                Ticker = Cells(rowcounter, 1)
                TickerClose = Cells(rowcounter, 6)
                
                TickerVolume = TickerVolume + Cells(rowcounter, 7)
                
                Cells(rowDisplay, 9) = Ticker
                Cells(rowDisplay, 10) = TickerClose - TickerOpen
                Cells(rowDisplay, 11) = (TickerClose - TickerOpen) / TickerOpen
                Cells(rowDisplay, 12) = TickerVolume
                
                rowDisplay = rowDisplay + 1
                
                ' ---------- DETERMINE GREATEST INCREASE, DECREASE AND TOTAL VOLUME --------------------------------------'
                If (TickerClose - TickerOpen) / TickerOpen > TickerGreatestIncreaseValue Then
                    TickerGreatestIncrease = Cells(rowcounter, 1)
                    TickerGreatestIncreaseValue = (TickerClose - TickerOpen) / TickerOpen
                End If
                
                If (TickerClose - TickerOpen) / TickerOpen < TickerGreatestDecreaseValue Then
                    TickerGreatestDecrease = Cells(rowcounter, 1)
                    TickerGreatestDecreaseValue = (TickerClose - TickerOpen) / TickerOpen
                End If
                
                If TickerVolume > TickerGreatestVolumeValue Then
                    TickerGreatestVolume = Cells(rowcounter, 1)
                    TickerGreatestVolumeValue = TickerVolume
                End If
                ' ---------- DETERMINE GREATEST INCREASE, DECREASE AND TOTAL VOLUME --------------------------------------'
            
                TickerOpen = Cells(rowcounter + 1, 3)   ' Get Open Value for new Ticker
                
                TickerVolume = 0    ' Reset volume to 0
                
            End If
            
            TickerVolume = TickerVolume + Cells(rowcounter, 7)
        
        Next rowcounter
        
        Range("K:K,Q2:Q3").NumberFormat = "0.00%"
        
        'Display Greatest _________ Information & Titles
        Cells(1, 16) = "Ticker": Cells(1, 17) = "Value"
        
        Cells(2, 15) = "Greatest % Increase"
        Cells(2, 16) = TickerGreatestIncrease
        Cells(2, 17) = TickerGreatestIncreaseValue
        
        Cells(3, 15) = "Greatest % Decrease"
        Cells(3, 16) = TickerGreatestDecrease
        Cells(3, 17) = TickerGreatestDecreaseValue
        
        Cells(4, 15) = "Greatest Total Volume"
        Cells(4, 16) = TickerGreatestVolume
        Cells(4, 17) = TickerGreatestVolumeValue
        
        ' Conditional Formatting -------------------------------------------------------------------'
        '   Highlight cells < 0 = RED...cells > 0 GREEN
        LastRow = Range("J" & Rows.Count).End(xlUp).Row     ' Last row in column J
        Dim myRange As Range
        Set myRange = Range("J2:J" & LastRow)
            
        myRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0"
        myRange.FormatConditions(1).Interior.ColorIndex = 10
           
        myRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
        myRange.FormatConditions(2).Interior.ColorIndex = 3
        ' Conditional Formatting -------------------------------------------------------------------'
    
        Cells.EntireColumn.AutoFit
    Next ws
    
    Application.ScreenUpdating = True ' Reset screen updating to default
End Sub



