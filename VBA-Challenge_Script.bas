Attribute VB_Name = "Module2"
Sub stock_market():
Dim ws As Worksheet

For Each ws In ThisWorkbook.Worksheets



Dim count As Integer
Dim first_value As Double
Dim last_value As Double
Dim stock_volume As LongLong
Dim yearly_change As Double
Dim first_found As Boolean

first_found = False
yearly_change = 0
stock_volume = 0
first_value = 0
count = 1

ws.Cells(1, 11).Value = "Ticker"
ws.Cells(1, 12).Value = "Yearly Change"
ws.Cells(1, 13).Value = "Percent Change"
ws.Cells(1, 14).Value = "Total Stock Volume"








For I = 2 To ws.Range("A" & Rows.count).End(xlUp).Row

    If first_found = True And ws.Cells(I, 1).Value <> ws.Cells(I + 1, 1).Value Then
        last_value = ws.Cells(I, 6).Value
        yearly_change = last_value - first_value
        ws.Cells(count + 1, 12).Value = yearly_change
        ws.Cells(count + 1, 13).Value = (last_value / first_value) - 1
        
    ElseIf first_found = False And ws.Cells(I, 1).Value = ws.Cells(I + 1, 1).Value Then
        first_found = True
        first_value = ws.Cells(I, 3).Value
        
    End If

    If ws.Cells(I, 1).Value = ws.Cells(I + 1, 1).Value Then
    stock_volume = stock_volume + ws.Cells(I, 7).Value
    
    Else
        ws.Cells(count + 1, 14).Value = stock_volume
        ws.Cells(count + 1, 11).Value = ws.Cells(I, 1).Value
        count = count + 1
        first_value = 0
        last_value = 0
        first_found = False
        stock_volume = 0
    
    End If
    
   
Next I

For I = 2 To ws.Range("L" & Rows.count).End(xlUp).Row
    If ws.Cells(I, 12) < 0 Then
        ws.Cells(I, 12).Interior.ColorIndex = 3
    Else
        ws.Cells(I, 12).Interior.ColorIndex = 4
        
    End If
    
Next I
    
Dim lastrow As Long
    lastrow = ws.Cells(Rows.count, "M").End(xlUp).Row


ws.Range("M2:M" & lastrow).NumberFormat = "0.00%"


ws.Cells(2, 16).Value = "Greatest % Increase"
ws.Cells(3, 16).Value = "Greatest % Decrease"
ws.Cells(4, 16).Value = "Greatest Total Volume"
ws.Cells(1, 17).Value = "Ticker"
ws.Cells(1, 18).Value = "Value"
Dim greatest As Double
Dim least As Double
Dim greatest_volume As LongLong
least = 0
greatest_volume = 0
greatest = 0
For I = 2 To ws.Range("M" & Rows.count).End(xlUp).Row
    If ws.Cells(I, 13).Value > greatest Then
        greatest = ws.Cells(I, 13).Value
        ws.Cells(2, 17).Value = ws.Cells(I, 11).Value
    End If
    
    If ws.Cells(I, 13).Value < least Then
        least = ws.Cells(I, 13).Value
        ws.Cells(3, 17).Value = ws.Cells(I, 11).Value
    End If
    
    If ws.Cells(I, 14).Value > greatest_volume Then
        greatest_volume = ws.Cells(I, 14).Value
       ws.Cells(4, 17).Value = ws.Cells(I, 11).Value
    End If
Next I

ws.Cells(2, 18).Value = greatest
ws.Cells(2, 18).NumberFormat = "0.00%"
ws.Cells(3, 18).Value = least
ws.Cells(3, 18).NumberFormat = "0.00%"
ws.Cells(4, 18).Value = greatest_volume


    
    
Next ws





End Sub

