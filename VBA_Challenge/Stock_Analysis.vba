Attribute VB_Name = "Module3"
Option Explicit
Dim i As Double
Dim j As Double
Dim k As Double
Dim tkr1 As String
Dim tkr2 As String
Dim tkr3 As String
Dim open_price As Single
Dim close_price As Single
Dim ychange As Single
Dim pchange As Single
Dim volume As Double
Dim checkincrease As Single
Dim checkdecrease As Single
Dim checkvol As Double
Dim current_greatincrease As Single
Dim current_greatdecrease As Single
Dim current_highestvol As Double
Dim sheetx As Worksheet
Dim LR As Double
Dim days As Integer



'All variable declaration made in 'Declarations' Section
'Run this sub!
Sub MultipleSheets()

    Application.ScreenUpdating = False
    
    'iterate through every sheet
    For Each sheetx In Worksheets
        sheetx.Select

        'run subs Stock_Analysis() and FormattCells on current sheet
        Call Stock_Analysis
        Call FormatCells
    Next
    
    Application.ScreenUpdating = True

End Sub


'All variable declaration made in 'Declarations' Section
Sub Stock_Analysis()
    
    'create headers for new tables
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"

    'set counter variables k and days to 0 and set LR equal to the last row that contains data
    k = 2
    days = 0
    LR = Cells(Rows.Count, 1).End(xlUp).Row
    
    'nested loop to determine how many trading days there are
    For i = 2 To LR
    
        tkr1 = Cells(i, 1).Value
        
        For j = 2 To LR
            days = days + 1
            tkr2 = Cells(j, 1).Value
            
            If tkr2 <> tkr1 Then
                days = days - 1
                i = LR
                Exit For
            End If
        Next j
    Next i
    
    'nested loop to evaluate each row and calculate output variables
    'outer loop runs for each stock, inner loop runs for each row within every stocks range(+1)
    For i = 2 To LR Step days

        volume = 0
        open_price = Cells(i, 3).Value
        tkr1 = Cells(i, 1).Value
        
        For j = i To LR

            volume = volume + Cells(j, 7).Value
            tkr2 = Cells(j, 1).Value

            'if line #97 is true you've hit the next ticker!
            If tkr2 <> tkr1 Then

                'subtract current cells volume in order to correct volume variable
                volume = volume - Cells(j, 7).Value

                'close price = previous rows close price value
                close_price = Cells(j - 1, 6).Value

                'make calculations and output results
                ychange = close_price - open_price
                pchange = ychange / open_price
                
                Cells(k, 9) = tkr1
                Cells(k, 10) = ychange
                Cells(k, 11) = pchange
                Cells(k, 12) = volume
                
                'format cells
                With Cells(k, 10).Interior
                    If Cells(k, 10).Value < 0 Then
                        .Color = RGB(255, 0, 0)
                    Else
                        .Color = RGB(0, 255, 0)
                    End If
                End With

                'add 1 to output row variable    
                k = k + 1
                'exit inner for loop in order to move onto next stock ticker
                Exit For
            End If
            
            'conditional to handle final stock ticker(j = last row)
            If j = LR Then
                
                close_price = Cells(j, 6).Value
                
                ychange = close_price - open_price
                pchange = ychange / open_price
                
                Cells(k, 9) = tkr1
                Cells(k, 10) = ychange
                Cells(k, 11) = pchange
                Cells(k, 12) = volume
                
                With Cells(k, 10).Interior
                
                    If Cells(k, 10).Value < 0 Then
                        .Color = RGB(255, 0, 0)
                    Else
                        .Color = RGB(0, 255, 0)
                    End If
                   
                End With
                Exit For
                
            End If
        Next j
    Next i
    
    'set comparison variables to 0 and LR equal to newly created tables last row
    current_greatincrease = 0
    current_greatdecrease = 0
    current_highestvol = 0
    LR = Cells(Rows.Count, 9).End(xlUp).Row
    
    'Loop to compare current variables to check variables
    For i = 2 To LR
        checkincrease = Cells(i, 11).Value
        checkdecrease = Cells(i, 11).Value
        checkvol = Cells(i, 12).Value
        
        If checkincrease > current_greatincrease Then
            current_greatincrease = checkincrease
            tkr1 = Cells(i, 9)
        End If
        
        If checkdecrease < current_greatdecrease Then
            current_greatdecrease = checkincrease
            tkr2 = Cells(i, 9)
        End If
        
        If checkvol > current_highestvol Then
            current_highestvol = checkvol
            tkr3 = Cells(i, 9)
        End If
    Next i
    
    'outputs values in new table
    Cells(2, 16).Value = tkr1
    Cells(2, 17).Value = current_greatincrease
    
    Cells(3, 16).Value = tkr2
    Cells(3, 17).Value = current_greatdecrease
    
    Cells(4, 16).Value = tkr3
    Cells(4, 17).Value = current_highestvol
    
End Sub

'sub for formatting cells in order to correctly display information output from Stock_Analysis()
Private Sub FormatCells()
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("J2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.NumberFormat = "0.00"
    Range("K2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Style = "Percent"
    Selection.NumberFormat = "0.00%"
    Range("Q2:Q3").Select
    Selection.Style = "Percent"
    Selection.NumberFormat = "0.00%"
    Range("Q4").Select
    Selection.NumberFormat = "0.00E+0"
End Sub

