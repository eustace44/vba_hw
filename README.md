

Sub stocktotal()
'Code to go through each worksheet tab
Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets

        ' Create a Variable to Hold File Name, Last Row, and Year
        Dim WorksheetName As String

        ' Determine the Last Row
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        ' Grabbed the WorksheetName
        WorksheetName = ws.Name
Dim i As Long
Dim Stock_Total As Double

Dim Yearly_Change As Long
Dim Counter As Long
Dim PriceFlag As Boolean
Dim percentMin As Double
Dim percentMax As Double
Dim volumeMax As Double
Dim percentMinTicker, percentMaxTicker, volumeMaxTicker As String
Counter = 2
PriceFlag = True
percentMin = 1E+99
percentMax = -1E+99
volumeMax = -1E+99
Stock_Total = 0


  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

  For i = 2 To lastrow

    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      Ticker_Name = ws.Cells(i, 1).Value

      Stock_Total = Stock_Total + ws.Cells(i, 7).Value

      closingPrice = ws.Cells(i, 6).Value
      
      Yearly_Change = closingPrice - openingPrice

        
      If Yearly_Change < 0 Then
                ws.Cells(Counter, 11).Interior.ColorIndex = 3
            ElseIf Yearly_Change > 0 Then
                ws.Cells(Counter, 11).Interior.ColorIndex = 4
            End If
            
      
       If Yearly_Change = 0 Or openingPrice = 0 Then
       Percent_Change = 0
       Else
       Percent_Change = Format(Yearly_Change / openingPrice, "#.##%")
        End If
        
    
    ws.Cells(Counter, 9).Value = Ticker_Name

      ws.Cells(Counter, 10).Value = Stock_Total
      
      ws.Cells(Counter, 11).Value = Yearly_Change

      ws.Cells(Counter, 12).Value = Percent_Change
       
If ws.Cells(Counter, 12).Value > percentMax Then
                If ws.Cells(Counter, 12).Value = ".%" Then
                Else
                    percentMax = ws.Cells(Counter, 12).Value
                    percentMaxTicker = Ticker_Name
                End If
            ElseIf ws.Cells(Counter, 12).Value < percentMin Then
                percentMin = ws.Cells(Counter, 12).Value
                percentMinTicker = Ticker_Name
            ElseIf Stock_Total > volumeMax Then
                volumeMax = Stock_Total
                volumeMaxTicker = Ticker_Name
                End If
      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Total
      Stock_Total = 0
      Counter = Counter + 1
      PriceFlag = True

    Else

      
      If PriceFlag Then
                openingPrice = ws.Cells(i, 3).Value
                PriceFlag = False
      
        End If
              ' Add to the Total
      Stock_Total = Stock_Total + ws.Cells(i, 7).Value

    End If

  Next i
  ws.Cells(2, 17).Value = Format(percentMax, "#.##%")
    ws.Cells(3, 17).Value = Format(percentMin, "#.##%")
    ws.Cells(4, 17).Value = volumeMax
    
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Total Stock Volume"
    ws.Cells(1, 11).Value = "Yearly Change"
    ws.Cells(1, 12).Value = "Percent Change"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Volume"
    ws.Cells(2, 16).Value = percentMaxTicker
    ws.Cells(3, 16).Value = percentMinTicker
    ws.Cells(4, 16).Value = volumeMaxTicker

 Next ws
End Sub
