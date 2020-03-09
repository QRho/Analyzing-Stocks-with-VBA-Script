Attribute VB_Name = "Module2"
Option Explicit
Sub Stocks()
' Set the initial variables
    
    Dim Ticker_Name As String
    Dim Change1 As Double
    Dim Monthly_Rating As Double
    Dim Stock_Volume As Variant
    Dim PERCENT_CHANGE As Double
    Dim Percent As Double
    Dim Stock_YearBegin As Double
    Dim Stock_YearEnd As Double
    Dim lastrow As Long
    Dim i As Long
    
'Zero out all counters

    Stock_YearBegin = 0
    Stock_YearEnd = 0
    Change1 = 0
    Stock_Volume = 0
    
'Create and set the Summary Table

    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
'Label Summary sheet
'I added both the opening and closeing values to the summary so that I could check them
'with the cooresponding data
    
    Range("I1").Select
    ActiveCell.FormulaR1C1 = "TickerName"
    Range("J1").Select
    ActiveCell.FormulaR1C1 = "Open"
    Range("K1").Select
    ActiveCell.FormulaR1C1 = "Close"
    Range("L1").Select
    ActiveCell.FormulaR1C1 = "Change"
    Range("M1").Select
    ActiveCell.FormulaR1C1 = "Percent"
    Range("N1").Select
    ActiveCell.FormulaR1C1 = "Volume"
    
    
' Grab the info you need before the loop
    
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    Stock_YearBegin = Cells(2, 3).Value
    Ticker_Name = Cells(2, 1).Value
    
    
' Set the location for the values you already have
    
    
    Cells(Summary_Table_Row, 9) = Ticker_Name
    Cells(Summary_Table_Row, 10) = Stock_YearBegin

'Begin For Loop
    
    For i = 2 To lastrow
              
        
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
            
'Defining & placing the Stock year end/beginning
            
            Stock_YearEnd = Cells(i, 6).Value
            
            Cells(Summary_Table_Row, 11) = Stock_YearEnd
            
            Change1 = Stock_YearEnd - Stock_YearBegin
            
' Grabbing the Stock Change and formatting it (considering the 0 value If, Then Procedure added*)
            
            Cells(Summary_Table_Row, 12) = Change1
            
            If Change1 = 0 Then
                Cells(Summary_Table_Row, 13) = 0
            Else
                Cells(Summary_Table_Row, 13) = Change1 / Stock_YearBegin
                Cells(Summary_Table_Row, 13).NumberFormat = "0.00%"
                
            End If
            
'Setting the color format for what is greater than and less 0
            
            If Cells(Summary_Table_Row, 13) > 0 Then
                
                Cells(Summary_Table_Row, 13).Interior.ColorIndex = 4
            Else
                Cells(Summary_Table_Row, 13).Interior.ColorIndex = 3
            End If
            
            Stock_YearBegin = Cells(i + 1, 6).Value
            
            
'Adding up and placing the Stock volume in Summary
            
            Cells(Summary_Table_Row, 14) = Stock_Volume
            
            Summary_Table_Row = Summary_Table_Row + 1
            
            Ticker_Name = Cells(i + 1, 1).Value
            
'Connecting Range to variable
            
            Range("N" & Summary_Table_Row).Value = Stock_Volume
            Range("I" & Summary_Table_Row).Value = Ticker_Name
            Range("J" & Summary_Table_Row).Value = Stock_YearBegin
            Range("K" & Summary_Table_Row).Value = Stock_YearEnd
            
'Zero'ing out Stock volume for last calculation
            
            Stock_Volume = 0
        Else
            

            Stock_Volume = Stock_Volume + Cells(i, 7).Value
            
            
            
            
        End If
        
        Next i
        
    End Sub

