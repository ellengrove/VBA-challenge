Attribute VB_Name = "Module1"
Sub stockAnalysis()

Dim ws As Worksheet

' Loop through each worksheet in workbook
For Each ws In ThisWorkbook.Worksheets

    ' Set headers for summary table
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    ' Declare variable that counts the row in the summary table and initialize at 2
    Dim rowCounter As Integer
    rowCounter = 2
    
    ' Declare variable that will store the first opening price per ticker; initializes at value in cell C2
    Dim firstOpen As Double
    firstOpen = ws.Cells(2, 3).Value
    
    ' Declare variable that will store the total yearly volume per ticker; initializes at value in cell G2
    Dim totalVolume As LongLong
    totalVolume = ws.Cells(2, 7).Value
    
    ' Declare variable that counts the last nonempty row of data
    Dim lastRow As Long
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Loop through all rows in spreadsheet
    For i = 2 To lastRow
        
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            ' Prints each ticker into the summary table
           ws.Cells(rowCounter, 9).Value = ws.Cells(i, 1).Value
            
            ' Calculates Yearly Change and Percent Change and and prints into the summary table
            ws.Cells(rowCounter, 10).Value = ws.Cells(i, 6).Value - firstOpen
            ws.Cells(rowCounter, 11).Value = ws.Cells(rowCounter, 10).Value / firstOpen
            
            ' Formats negative Yearly Change values with a red cell and positive Yearly Change values with a green cell
            If ws.Cells(rowCounter, 10).Value < 0 Then
                ws.Cells(rowCounter, 10).Interior.ColorIndex = 3
            ElseIf ws.Cells(rowCounter, 10).Value > 0 Then
                ws.Cells(rowCounter, 10).Interior.ColorIndex = 4
            End If
            
            ' Resets the first opening price for the next ticker in the list
            firstOpen = ws.Cells(i + 1, 3).Value
            
            ' Prints the yearly volume in the summary table and resets volume to the first day's volume for the next ticker
            ws.Cells(rowCounter, 12).Value = totalVolume
            totalVolume = ws.Cells(i + 1, 7).Value
                
            ' Sets a new row in the summary table
            rowCounter = rowCounter + 1
            
        Else
        
            ' Aggregates total yearly volume across each ticker
            totalVolume = totalVolume + ws.Cells(i + 1, 7).Value
           
        End If
            
    Next i
    
' Bonus summary table

    ' Set headers for summary table
    ws.Range("N2").Value = "Greatest % Increase"
    ws.Range("N3").Value = "Greatest % Decrease"
    ws.Range("N4").Value = "Greatest Total Volume"
    ws.Range("O1").Value = "Ticker"
    ws.Range("P1").Value = "Value"
    
    ' Declare variable that counts the last nonempty row of data in bonus summary table
    Dim bonusRow As Integer
    bonusRow = ws.Cells(Rows.Count, 11).End(xlUp).Row

    ' Set initial values in bonus summary table
    ws.Range("P2").Value = ws.Range("K2").Value
    ws.Range("P3").Value = ws.Range("K2").Value
    ws.Range("P4").Value = ws.Range("L2").Value
    ws.Range("O2").Value = ws.Range("I2").Value
    ws.Range("O3").Value = ws.Range("I2").Value
    ws.Range("O4").Value = ws.Range("I2").Value

    
    ' Loop through summary table
    For j = 2 To bonusRow
        
        ' Look for greatest % increase
        If ws.Cells(j, 11).Value > ws.Range("P2").Value Then
            ws.Range("O2").Value = ws.Cells(j, 9).Value
            ws.Range("P2").Value = ws.Cells(j, 11).Value
        End If
        
        ' Look for greatest % decrease
        If ws.Cells(j, 11).Value < ws.Range("P3").Value Then
            ws.Range("O3").Value = ws.Cells(j, 9).Value
            ws.Range("P3").Value = ws.Cells(j, 11).Value
        End If
        
        ' Look for greatest yearly trade volume
        If ws.Cells(j, 12).Value > ws.Range("P4").Value Then
            ws.Range("O4").Value = ws.Cells(j, 9).Value
            ws.Range("P4").Value = ws.Cells(j, 12).Value
        End If
        
    Next j
    
        ' Resize columns to accomodate large volume numbers
        ws.Columns(12).AutoFit
        ws.Columns(14).AutoFit
        ws.Columns(16).AutoFit
        
        ' Apply number formats
        ws.Range("J:J").NumberFormat = "0.00"
        ws.Range("K:K").NumberFormat = "0.00%"
        ws.Range("P2:P3").NumberFormat = "0.00%"
    
Next


End Sub

