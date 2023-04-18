Attribute VB_Name = "Module1"
Sub stock_analysis():
'Create for loop to run script on all sheets
'For Each ws In Worksheets
Dim i As Long
Dim total As Double
Dim change As Single
Dim j As Integer
Dim start As Long
Dim rowCount As Long
Dim percentChange As Single
Dim days As Integer
Dim dailyChange As Single
Dim averageChange As Single

'set summary table column labels
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"

'SET initial values
'Declare and set initial Summary_Table_Row to 2
'Dim Summary_Table_Row As Integer
'Summary_Table_Row = 2
j = 0
total = 0
start = 2
change = 0
'getting the row number to count the data , and it is the last row and Find length of columb "A" and store As Long
'Dim Rowcount As Long
rowCount = Cells(Rows.Count, "A").End(xlUp).Row

'Create for loop to detect change in Ticker
For i = 2 To rowCount

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        'Store <ticker> for cells that satisfy condition above as Ticker
        'Ticker = ws.Cells(i, 1).Value
        'and store results
        total = total + Cells(i, 7).Value
        
        'store the volume in the variable as zero
        If total = 0 Then
            ' print the results
            Range("I" & 2 + j).Value = Cells(i, 1).Value
            Range("J" & 2 + j).Value = 0
            Range("K" & 2 + j).Value = "%" & 0
            Range("L" & 2 + j).Value = 0
        Else
        ' find first non zero value
        If Cells(start, 3) = 0 Then
            For Find_Value = start To i
                If Cells(Find_Value, 3).Value <> 0 Then
                Exit For
            End If
        Next Find_Value
    End If
        ' calculating the change
        
        change = (Cells(i, 6) - Cells(start, 3))
        PercentCange = change / Cells(start, 3)
        'setting the start count of the next ticker as
        start = i + 1
        ' printing the results
        Range("I" & 2 + j).Value = Cells(i, 1).Value
        Range("J" & 2 + j).Value = change
        Range("J" & 2 + j).NumberFormat = "0.00%"
        Range("K" & 2 + j).NumberFormat = percentChange
        Range("K" & 2 + j).NumberFormat = " 0.00%"
        Range("L" & 2 + j).Value = total
        
        'selecting the color code for positive value with green and negative value in the cell with the red color
        'selecting the location to change the color
        Select Case change
            Case Is > 0
            Range("J" & 2 + j).Interior.ColorIndex = 4
             Case Is < 0
            Range("J" & 2 + j).Interior.ColorIndex = 3
             Case Else
            Range("J" & 2 + j).Interior.ColorIndex = 0
        End Select
    End If
        'Reset/adjust necessary variables
        total = 0
        change = 0
        j = j + 1
        days = 0
    'If ticker is same add results
    Else
        total = total + Cells(i, 7).Value
    End If
    Next i
End Sub
