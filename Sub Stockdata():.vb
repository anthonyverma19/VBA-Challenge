Sub Stockdata():
'select Worksheet 1
Dim CurrentWs As Worksheet
'store results table and headers true and false
Dim Results_Sheet As Boolean
Need_Summary_Table_Header = True

'loop data through all worksheets
For Each CurrentWs In Worksheets

'create variable for "ticker names"
Dim Ticker_Name As String

'create variable for volume
Dim Total_Volume As Double
Total_Volume = 0

' create variable for opening and closing prices
'create variable for yearly change and percent
Dim Opening_Price As Double
Opening_Price = 0
Dim Closing_Price As Double
Closing_Price = 0
Dim Yearly_Change As Double
Yearly_Change = 0
Dim Yearly_Percent As Double
Yearly_Percent = 0

' create variables for bonus
Dim Bonus_Increase As Double
Bonus_Increase = 0
Dim Bonus_Decrease As Double
Bonus_Decrease = 0
Dim Greatest_Volume As Double
Greatest_Volume = 0

' create variable tickers for bonus values

Dim Bonus_Increase_Ticker As String
Dim Bonus_Decrease_Ticker As String
Dim Greatest_Volume_Ticker As String

'create a location for the tickers
Dim Summary_Table_Row As Long
Summary_Table_Row = 2

'set variable for the last row
Dim LastRow As Long
Dim i As Long


'retrieve LastRow from every worksheet
LastRow = CurrentWs.Cells(Rows.Count, 1).End(xlUp).Row

'create headers for results table
If Need_Summary_Table_Header Then
CurrentWs.Range("I1").Value = "Ticker"
CurrentWs.Range("J1").Value = "Yearly Change"
CurrentWs.Range("K1").Value = "Percent Change"
CurrentWs.Range("L1").Value = "Total Stock Volume"

'create headers for bonus values
CurrentWs.Range("O2").Value = "Greatest % Increase"
CurrentWs.Range("o3").Value = "Greatest % Decrease"
CurrentWs.Range("O4").Value = "Highest Volume"
CurrentWs.Range("P1").Value = "Ticker"
CurrentWs.Range("Q1").Value = "Value"

Else
Need_Summary_Table_Header = True
End If

Opening_Price = CurrentWs.Cells(2, 3).Value



'loop through and seperate different stocks with tickers

For i = 2 To LastRow

'set if statement for the change in the ticker name

If CurrentWs.Cells(i + 1, 1).Value <> CurrentWs.Cells(i, 1).Value Then
Ticker_Name = CurrentWs.Cells(i, 1).Value

'Calculating yearly change and yearly percent

Closing_Price = CurrentWs.Cells(i, 6).Value
Yearly_Change = (Closing_Price - Opening_Price)

'Set if statement for yearly change and yearly percent
If Opening_Price <> 0 Then
Yearly_Percent = (Yearly_Change / Opening_Price) * 100
End If

'add ticker name to total volume

Total_Volume = Total_Volume + CurrentWs.Cells(i, 7).Value

'put ticker name, volume, yearly change, and yearly percent in the summary tables
CurrentWs.Range("I" & Summary_Table_Row).Value = Ticker_Name
CurrentWs.Range("L" & Summary_Table_Row).Value = Total_Volume
CurrentWs.Range("J" & Summary_Table_Row).Value = Yearly_Change

' next insert percent sign
CurrentWs.Range("K" & Summary_Table_Row).Value = (CStr(Yearly_Percent) & "%")

'color code columns
If (Yearly_Change > 0) Then
    CurrentWs.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
ElseIf (Yearly_Change <= 0) Then
CurrentWs.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
End If

'row count plus one

Summary_Table_Row = Summary_Table_Row + 1


'bonus calculations
Bonus_Increase = WorksheetFunction.Max(CurrentWs.Range("K:K"))
Bonus_Decrease = WorksheetFunction.Min(CurrentWs.Range("K:K"))
Greatest_Volume = WorksheetFunction.Max(CurrentWs.Range("L:L"))


'put bonus values in a summary table

CurrentWs.Range("Q2").Value = (CStr(Bonus_Increase) & "%")
CurrentWs.Range("Q3").Value = (CStr(Bonus_Increase) & "%")
CurrentWs.Range("P2").Value = Bonus_Increase_Ticker
CurrentWs.Range("P3").Value = Bonus_Decrease_Ticker
CurrentWs.Range("Q4").Value = Greatest_Volume
CurrentWs.Range("P4").Value = Greatest_Volume_Ticker

'reset ticker values
Yearly_Percent = 0
Total_Volume = 0
Yearly_Change = 0
Closing_Price = 0

'recreate opening price
Opening_Price = CurrentWs.Cells(i + 1, 3).Value

'add stock volume cells to stock volume
Else
Total_Volume = Total_Volume + CurrentWs.Cells(i, 7).Value


End If

Next i


Next CurrentWs

End Sub
