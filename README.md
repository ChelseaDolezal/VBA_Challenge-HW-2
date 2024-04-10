# VBA_Challenge-HW-2
Name VBA Code
Sub Chelsea()


'Define Variables
Dim Ticker As String
Dim TotalVolume As Double
Dim Output_Table_Row As Integer
Dim Row As Integer
Dim YearlyChange As Double
Dim PercentChange As Double
Dim Opening As Double
Dim Closing As Double
Dim ws As Worksheet
Dim i As Integer
Dim j As Integer
Dim LastRow As Integer
Dim GreatestIncrease_ticker As String
Dim GreatestIncrease_value As Double
Dim GreatestDecrease_ticker As String
Dim GreatestDecrease_value As Double
Dim GreatestTotalVolme_ticker As String
Dim GreatestTotalVolume_value As Double


'loop all worksheets in workbook
For Each ws In ThisWorkbook.Worksheets

'activate worksheets
ws.Activate

'set starting values
TotalVolume = 0
Output_Table_Row = 2
i = 2
j = 3

'define last row
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'loop rows
For Row = 2 To LastRow

'find when ticker name changes
If Cells(Row + 1, 1).Value <> Cells(Row, 1).Value Then

'calculate/store desired data
    Ticker = Cells(Row, 1).Value
    TotalVolume = TotalVolume + Cells(Row, 7).Value
    YearlyChange = Cells(Row, 6).Value - Cells(i, j).Value
    PercentChange = (YearlyChange / Cells(i, j).Value) * 100

    i = Row + 1

'print data into new table
    Range("I" & Output_Table_Row).Value = Ticker
    Range("L" & Output_Table_Row).Value = TotalVolume
    Range("J" & Output_Table_Row).Value = YearlyChange
    Range("K" & Output_Table_Row).Value = PercentChange

'move down one table row and reset total volume for next ticker
    Output_Table_Row = Output_Table_Row + 1
    TotalVolume = 0

'calculate total volume
Else
    TotalVolume = TotalVolume + Cells(Row, 7).Value

End If

'loop to next row
Next Row


'redefine last row for created data table
LastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row

'set values to "0"
GreatestIncrease_value = 0
GreatestDecrease_value = 0
GreatestTotalVolume_value = 0


'loop for changing color cells based on yearly change and percent change
For Row = 2 To LastRow

'change color of cells for yearly change
If Cells(Row, 10).Value >= 0 Then
    Cells(Row, 10).Interior.ColorIndex = 4

Else
    Cells(Row, 10).Interior.ColorIndex = 3

End If

'change color of cells for percent change
If Cells(Row, 11).Value >= 0 Then
    Cells(Row, 11).Interior.ColorIndex = 4

Else
    Cells(Row, 11).Interior.ColorIndex = 3

End If

'loop to next row
Next Row



'redifine new output table start
Output_Table_Row = 2

'loop for greatest % increase
For i = 2 To LastRow

'find/store greatest % increase data and corresponding ticker name
    If Cells(i, 11) > GreatestIncrease_value Then
        GreatestIncrease_value = Cells(i, 11)
        Ticker = Cells(i, 9)
        
'print greatest % increase and ticker in new data table
        Range("O" & Output_Table_Row).Value = Ticker
        Range("P" & Output_Table_Row).Value = GreatestIncrease_value
        
    End If
    
'loop next row
Next i

'define next table row
Output_Table_Row = 3

'loop for greatest % decrease
For i = 2 To LastRow

'find/store greatest % decrease data and corresponding ticker name
    If Cells(i, 11) < GreatestDecrease_value Then
        GreatestDecrease_value = Cells(i, 11)
        Ticker = Cells(i, 9)
        
'print greatest % decrease and ticker in data table
        Range("O" & Output_Table_Row).Value = Ticker
        Range("P" & Output_Table_Row).Value = GreatestDecrease_value
        
    End If
    
'loop next row
Next i

'define next table row
Output_Table_Row = 4

'loop for greatest total volume
For i = 2 To LastRow

'find/store greatest total volume and corresponding ticker name
    If Cells(i, 11) > GreatestTotalVolume_value Then
        GreatestTotalVolume_value = Cells(i, 12)
        Ticker = Cells(i, 9)
        
'print greatest total volume and ticker in data table
        Range("O" & Output_Table_Row).Value = Ticker
        Range("P" & Output_Table_Row).Value = GreatestTotalVolume_value
        
    End If
'loop next row
Next i

'loop next worksheet
Next ws


End Sub
