Attribute VB_Name = "Module1"
Sub Stock()
'create variables to hold the worksheet, worksheet name
Dim ws As Worksheet
Dim WorksheetName As Integer
WorksheetName = Name

'create a variable c to hold the column number which is intialy set to 250 and changes for each sheet.
'This is used to calculate the opening price for each ticker.
'Create a variable to hold the last row.
Dim c As Long
Dim LastRow As Long

'Loop through all sheets
For Each ws In ThisWorkbook.Worksheets
If ws.Name = "2018" Then
    c = 250
    ElseIf ws.Name = "2019" Then
    c = 251
    ElseIf ws.Name = "2020" Then
    c = 252
End If

'Determine the LastRow
LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row

'Display the row header as mentioned in the challenge
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

'Create variables to hold the values and perform the calculations
Dim Yearly_Change As Double
Dim Volume As Double
Dim Opening_Price As Double
Dim Closing_Price As Double
Dim Percentage_change As Double
Dim row As Integer
row = 2
Volume = 0

'Create Object to store the range of values
Dim ColumnRange As Range
Dim TotalVolume As Range


For i = 2 To LastRow
    
    If (ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value) Then
        
        ws.Cells(row, 9).Value = ws.Cells(i, 1).Value
        Opening_Price = ws.Cells(i - c, 3).Value
        Closing_Price = ws.Cells(i, 6).Value
        ws.Cells(row, 10).Value = Closing_Price - Opening_Price
        Yearly_Change = ws.Cells(row, 10).Value
         
         'Color the cells red or green based on the values negative or postive respectively
         If (Yearly_Change > 0) Then
             ws.Cells(row, 10).Interior.ColorIndex = 4
         Else
            ws.Cells(row, 10).Interior.ColorIndex = 3
         End If
        
        Percentage_change = Round((Yearly_Change / Opening_Price) * 100, 2)
        ws.Cells(row, 11).Value = Format(Percentage_change, "0.00") & "%"
        
        'Color the cells red or green based on the values negative or postive respectively
        If (ws.Cells(row, 11).Value > 0) Then
             ws.Cells(row, 11).Interior.ColorIndex = 4
         Else
            ws.Cells(row, 11).Interior.ColorIndex = 3
         End If
         
        Volume = Volume + ws.Cells(i, 7).Value
        ws.Cells(row, 12).Value = Volume
        
        'Reset the volume
        Volume = 0
        row = row + 1
    Else
       'If the cell immediately following a row is the same ticker symbol
        Volume = Volume + ws.Cells(i, 7).Value
        
    End If
    
Next i


Set ColumnRange = ws.Range("K:K")
Set TotalVolume = ws.Range("L:L")


ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"

'Determine the maximum and minimum percent change
'Determine the highest total volume
ws.Range("Q2").Value = FormatPercent(Application.WorksheetFunction.Max(ColumnRange))
ws.Range("Q3").Value = FormatPercent(Application.WorksheetFunction.Min(ColumnRange))
ws.Range("Q4").Value = Application.WorksheetFunction.Max(TotalVolume)
ws.Range("Q4").Value = FormatNumber(ws.Range("Q4").Value)

'Loop through the sheets to display the maximum and minimum value
For i = 2 To 3001
    If (ws.Range("Q2").Value = ws.Cells(i, 11).Value) Then
        ws.Range("P2").Value = ws.Cells(i, 9).Value
    End If
    If (ws.Range("Q3").Value = ws.Cells(i, 11).Value) Then
        ws.Range("P3").Value = ws.Cells(i, 9).Value
    End If
    If (ws.Range("Q4").Value = ws.Cells(i, 12).Value) Then
        ws.Range("P4").Value = ws.Cells(i, 9).Value
    End If
Next i
Next ws

'msg box to display the task completed
  MsgBox ("Task Completed")
    
End Sub
