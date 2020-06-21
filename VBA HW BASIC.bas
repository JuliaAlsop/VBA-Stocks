Sub basic_part()
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Range("I1:L1").Columns.AutoFit
   
  ' Set an initial variable for holding the ticker name
  Dim Ticker_Name As String

' Set an initial variable for holding the open price per ticker name
  Dim OpenPrice_Total As Double
  OpenPrice_Total = 0

  ' Set an initial variable for holding the close price per ticker name
  Dim ClosePrice_Total As Double
  ClosePrice_Total = 0

  ' Set an initial variable for holding the total yearly change per ticker name
  Dim YearlyChange_Total As Double
  YearlyChange_Total = 0

  ' Set an initial variable for holding the total percent change per ticker name
  Dim PercentChange_Total As Double
  PercentChange_Total = 0

  ' Set an initial variable for holding the total volume per ticker name
  Dim Volume_Total As Double
  Volume_Total = 0

  ' Keep track of the location for each ticker name in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

     ' Loop through all ticker volumes
  For i = 2 To 70926

      ' Check if we are still within the same ticker name, if it is not...
      If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Set the Ticker name
      Ticker_Name = Cells(i, 1).Value
      
      'To retrieve the last Close Price Value per Ticker Name
      
     
      'Set Close Price Value
     ClosePrice_Total = YearlyChange_Total + Cells(i, 3).Value
      
      'To retrieve First Open Price Value per Ticker Name
      
      
        'Set Open Price Value
      OpenPrice_Total = PercentChange_Total + Cells(i, 6).Value

      ' Add to the Yearly Change Total
      YearlyChange_Total = OpenPrice_Total - ClosePrice_Total

      ' Add to the Percent Change Total
       PercentChange_Total = (ClosePrice_Total / OpenPrice_Total)

      ' Add to the Volume Total
      Volume_Total = Volume_Total + Cells(i, 7).Value

      ' Print the Ticker Name in the Summary Table
      Range("I" & Summary_Table_Row).Value = Ticker_Name

      ' Print the Yearly Change Amount to the Summary Table
      Range("J" & Summary_Table_Row).Value = YearlyChange_Total

      ' Print the Percent Change Amount to the Summary Table
      Range("K" & Summary_Table_Row).Value = PercentChange_Total
      
      ' Print the Volume Amount to the Summary Table
      Range("L" & Summary_Table_Row).Value = Volume_Total

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1

      ' Reset the Open Price Total
      OpenPrice_Total = 0

      ' Reset the Close Price Total
      ClosePrice_Total = 0

      ' Reset the Yearly Change Total
      YearlyChange_Total = 0
      
      ' Reset the Percent Change Total
      PercentChange_Total = 0
      
      ' Reset the Volume Total
      Volume_Total = 0

     ' If the cell immediately following a row is the same brand...
     Else

      ' Add to the Volume Total
      Volume_Total = Volume_Total + Cells(i, 7).Value

     End If
     
  Next i
  'Conditional Formatting Attempt
 'Dim j As Long, r1 As Range, r2 As Range

   'For j = 2 To 80000
      'Set r1 = Range("J" & j)
      'Set r2 = Range("J" & j)
      'If r1.Value >= 1 Then r1.Interior.Color = vbGreen
      'If r2.Value <= -1 Then r2.Interior.Color = vbRed
   'Next j
   'Variation of stackoverflow example:https://stackoverflow.com/questions/35142985/vba-change-color-of-cells-based-on-value-in-particular-cell
End Sub
'Worksheet Loop Attempt
'Dim WS_Count As Integer
    'Dim H As Integer
    'WS_Count = ActiveWorkbook.Worksheets.Count
    'Begin of Worksheet loop
    'For H = 1 To WS_Count
    'Next H
 
 'Conditional Formatting Attempt
 'Dim j As Long, r1 As Range, r2 As Range

   'For j = 2 To LastRow
      'Set r1 = Range("J1:LastRow" & j)
      'Set r2 = Range("J1:LastRow" & j)
      'If r1.Value >= 0 Then r1.Interior.Color = vbGreen
      'If r2.Value <= -1 Then r2.Interior.Color = vbRed
   'Next j
   'Variation of stackoverflow example:https://stackoverflow.com/questions/35142985/vba-change-color-of-cells-based-on-value-in-particular-cell



