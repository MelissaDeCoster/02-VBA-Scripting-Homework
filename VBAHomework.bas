Attribute VB_Name = "VBAHomework"
Sub CreateStockSummary()


' Set an initial variables
Dim ticker As String
Dim opening As Double
Dim closing As Double
Dim change As Double
Dim changepct As Double
Dim volume As Double
Dim MaxRow As Long
Dim k As Integer
Dim SheetCount As Long

Application.ScreenUpdating = False

SheetCount = ActiveWorkbook.Worksheets.Count

    'Loop Though each sheet
    
    
    For k = 1 To SheetCount
    Sheets(k).Select
        
        MaxRow = Range("A" & Rows.Count).End(xlUp).Row

        volume = 0

    'Enter Summary Table Headers
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"

    ' Keep track of the location for each ticker in the summary table
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2

    'Set the Opening Price

        opening = Cells(2, 3).Value

      ' Loop through all stock data
          For i = 2 To MaxRow


            ' Check if we are still within the same credit card brand, if it is not...
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

              ' Set the Brand name
              ticker = Cells(i, 1).Value

              ' Add to the Brand Total
              volume = volume + Cells(i, 7).Value

              'Set the Closing Price
              closing = Cells(i, 6).Value

              'Calculate the Change
              change = closing - opening

              'Calculate the Percent Change
                If change = 0 Or opening = 0 Then
                changepct = 0
                Else
                changepct = change / opening
                End If

              ' Print the Stock Ticker in the Summary Table
              Range("I" & Summary_Table_Row).Value = ticker

              'Print the Yearly Change to the Summary Table and apply conditional formatting
              Range("J" & Summary_Table_Row).Value = change
                If change <= 0 Then
                Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                Else
                Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                End If

              'Print the Percent Change to the Summary Table
              Range("K" & Summary_Table_Row).Value = Format(changepct, "0.00%")

              ' Print the Stock Volume to the Summary Table
              Range("L" & Summary_Table_Row).Value = volume

              ' Add one to the summary table row
              Summary_Table_Row = Summary_Table_Row + 1

              ' Reset the Brand Total
              volume = 0

              'Reset the Opening
              opening = Cells(i + 1, 3).Value


            ' If the cell immediately following a row is the same brand...
            Else

              ' Add to the Brand Total
              volume = volume + Cells(i, 7).Value

            End If

          Next i
          
            'Set up second summary table
              Range("P1").Value = "Ticker"
              Range("Q1").Value = "Amount"
            
              MaxRow = Range("I" & Rows.Count).End(xlUp).Row
              
              pctincr = 0
              pctdecr = 0
              vol2 = 0
            
            'Loop through summary to obtain values for second summary
            For y = 2 To MaxRow
            
            'Determine which ticker as the greatest % inc
                If Cells(y, 11).Value > pctincr Then
                pctincr = Cells(y, 11).Value
                pctincticker = Cells(y, 9).Value
                Else
                End If
                
            'Determine which ticker has the greatest volume
                If Cells(y, 12).Value > vol2 Then
                vol2 = Cells(y, 12).Value
                vol2ticker = Cells(y, 9).Value
                Else
                End If
                
            'Determine which ticker as the greatest % dec
                If Cells(y, 11).Value < pctdecr Then
                pctdecr = Cells(y, 11).Value
                pctdecticker = Cells(y, 9).Value
                Else
                End If
                
                Next y
                
          
            ' Print the Greatest Percent Increase in the Summary Table
              Range("O2").Value = "Greatest % Increase"
              Range("P2").Value = pctincticker
              Range("Q2").Value = Format(pctincr, "0.00%")
              
            ' Print the Greatest Percent Decrease in the Summary Table
              Range("O3").Value = "Greatest % Decrease"
              Range("P3").Value = pctdecticker
              Range("Q3").Value = Format(pctdecr, "0.00%")
    
            ' Print the Greatest Total Volume in the Summary Table
              Range("O4").Value = "Greatest Total Volume"
              Range("P4").Value = vol2ticker
              Range("Q4").Value = vol2
    
    'Auto Fit Summary Column-width so they look nice
    Columns("I:L").EntireColumn.AutoFit
    Columns("O:Q").EntireColumn.AutoFit
    
    Range("A1").Select
    
    Next k

Sheets(1).Select

Application.ScreenUpdating = True

End Sub


Sub ClearStockSummary()

SheetCount As Integer

For i = 1 To SheetCount

    Sheets(i).Select
    Columns("I:Q").ClearContents
    Columns("I:Q").Interior.ColorIndex = 0
    Columns("I:Q").ColumnWidth = 10.08
   
Next i
    
End Sub



