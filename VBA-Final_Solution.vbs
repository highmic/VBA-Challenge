Sub StockAnalysis_Tunde()

'Declare ws as worksheet variable in the "For Each" loop to ensure application to each worksheet in the workbook
Dim ws As Worksheet
For Each ws In Worksheets
 
'Basic solution variables declarations 
 Dim TickerSymbol As String
 Dim Summary_Table_Row As Integer
 Dim YearlyChange As Double
 Dim OpenPrice As Double
 Dim ClosePrice As Double
 Dim PercentChange As Double
 Dim TotalStkVol As Double

 'Assign initial value to basic variables 
 Summary_Table_Row = 2
 TotalStkVol = 0
 OpenPrice = ws.Cells(2, 3).Value
 ClosePrice = 0
 YearlyChange = 0
 PercentChange = 0
 
  'I need to ensure loop goes to the end of rows with data in column 1,use the End(xlDown)function
 
  LastRow = ws.Cells(2, 1).End(xlDown).Row
 
 'To ensure output printed on each worksheets, add ws before Cells, else code will print to active sheet only
 'Add headers for summary columns in each worksheet for the basic solution
    ws.Cells(1, 11).Value = "Ticker"
    ws.Cells(1, 12).Value = "Yealy Change"
    ws.Cells(1, 13).Value = "Percent Change"
    ws.Cells(1, 14).Value = "Total Stock Volume"

 'Add headers for summary columns in each worksheet for the challenge solution
     ws.Cells(2, 16).Value = "Greatest % Increase"
     ws.Cells(3, 16).Value = "Greatest % Decrease"
     ws.Cells(4, 16).Value = "Greatest Total Volume"
     ws.Cells(1, 17).Value = "Ticker"
     ws.Cells(1, 18).Value = "Value"
 
 For I = 2 To LastRow
    
   If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
     TickerSymbol = ws.Cells(I, 1).Value
     TotalStkVol = TotalStkVol + ws.Cells(I, 7).Value
       ClosePrice = ws.Cells(I, 6).Value
       YearlyChange = ClosePrice - OpenPrice

      'Add conditions for the divisor "Open price" to prevent runtime errors 11 and 6
         If OpenPrice <> 0 Then
            PercentChange = (YearlyChange / OpenPrice) * 100
         Else
           PercentChange = 100
           
      End If
    'Print outputs in the specified columns and ensure it begins after summary table row
     ws.Range("K" & Summary_Table_Row).Value = TickerSymbol
     ws.Range("N" & Summary_Table_Row).Value = TotalStkVol
     ws.Range("L" & Summary_Table_Row).Value = YearlyChange
     ws.Range("M" & Summary_Table_Row).Value = PercentChange & "%"

     'Add conditions for background colors for postive and negative stock increases
    If YearlyChange > 0 Then
      ws.Range("L" & Summary_Table_Row).Interior.ColorIndex = 4
      
    Else
     ws.Range("L" & Summary_Table_Row).Interior.ColorIndex = 3
     
    End If
    'Add 1 to summary row to prevent next ticker overwrite
     Summary_Table_Row = Summary_Table_Row + 1

    'Reset value of variables to initial value for next Ticker
    TotalStkVol = 0
    YearlyChange = 0
    PercentChange = 0
    
    ClosePrice = 0
    OpenPrice = ws.Cells(I + 1, 3).Value
     
    Else
    
      TotalStkVol = TotalStkVol + ws.Cells(I, 7).Value
    
   
    End If
  Next I

 'Challenge solution

 'Declare variables for challenge solution
 Dim GreatestPerIncrease As Double
 Dim GreatestPerDecrease As Double
 Dim GreatestTotVol As Double

 'Assign values to challenge variables, use VBA function to capture max/min values in a column

 GreatestPerIncrease = Application.WorksheetFunction.Max(ws.Columns("M"))
 GreatestPerDecrease = Application.WorksheetFunction.Min(ws.Columns("M"))
 GreatestTotVol = Application.WorksheetFunction.Max(ws.Columns("N"))
 LastRowPercent = ws.Cells(2, 13).End(xlDown).Row
 LastRowVolume = ws.Cells(2, 14).End(xlDown).Row

   For j = 2 To LastRowPercent
  If GreatestPerIncrease = ws.Cells(j, 13).Value Then
  
    ws.Cells(2, 18).Value = ws.Cells(j, 13).Value
    ws.Cells(2, 17).Value = ws.Cells(j, 11).Value
    ws.Cells(2, 18).Style = "Percent"
    ElseIf GreatestPerDecrease = ws.Cells(j, 13).Value Then
  
    ws.Cells(3, 18).Value = ws.Cells(j, 13).Value
    ws.Cells(3, 17).Value = ws.Cells(j, 11).Value
    ws.Cells(3, 18).Style = "Percent"
  End If

Next j

 For k = 2 To LastRowVolume
  If GreatestTotVol = ws.Cells(k, 14).Value Then
  
    ws.Cells(4, 18).Value = ws.Cells(k, 14).Value
    ws.Cells(4, 17).Value = ws.Cells(k, 11).Value
    
 End If
 Next k
   'Repeat in the next worksheet
 Next ws

End Sub