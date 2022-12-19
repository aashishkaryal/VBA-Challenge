{\rtf1\ansi\ansicpg1252\cocoartf2706
\cocoatextscaling0\cocoaplatform0{\fonttbl\f0\fnil\fcharset0 HelveticaNeue;}
{\colortbl;\red255\green255\blue255;}
{\*\expandedcolortbl;;}
\margl1440\margr1440\vieww11520\viewh8400\viewkind0
\deftab560
\pard\pardeftab560\slleading20\pardirnatural\partightenfactor0

\f0\fs26 \cf0 Sub multiple_year_stock_data()\
\
'defining worksheet\
Dim ws As Worksheet\
\
'defining ticker\
Dim ticker As String\
\
'defining volume\
Dim vol As Double\
\
'defining for prices and changes\
Dim open_price As Double\
Dim close_price As Double\
Dim yearly_change As Double\
Dim percent_change As Double\
\
'defining to format color\
Dim rg As Range\
Dim g As Long\
Dim c As Long\
Dim color_cell As Range\
\
'running through all worksheet\
For Each ws In ThisWorkbook.Worksheets\
\
'creating column header\
ws.Cells(1, 9).Value = "Ticker"\
ws.Cells(1, 10).Value = "Yearly Change"\
ws.Cells(1, 11).Value = "Percent Change"\
ws.Cells(1, 12).Value = "Total Stock Volume"\
Range("O1").Value = "Ticker"\
Range("P1").Value = "Value"\
Range("N2").Value = " Greatest % Increase"\
Range("N3").Value = " Greatest % Decrease"\
Range("N4").Value = " Greatest Total Volume"\
\
'setting up integers for loop\
Summary_Table_Row = 2\
\
'looping\
\
For I = 2 To ws.UsedRange.Rows.Count\
If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then\
\
'finding all the values\
\
ticker = ws.Cells(I, 1).Value\
vol = ws.Cells(I, 7).Value\
open_price = ws.Cells(I, 3).Value\
close_price = ws.Cells(I, 6).Value\
yearly_change = close_price - open_price\
percent_change = (close_price - open_price) / close_price\
\
End If\
Next I\
\
'setting column K as percent\
ws.Columns("K").NumberFormat = "0.00%"\
\
\
'setting the range\
Set rg = ws.Range("J2", Range("J2").End(xlDown))\
c = rg.Cells.Count\
\
'formatting color\
For g = 1 To c\
Set color_cell = rg(g)\
Select Case color_cell\
\
Case Is >= 0\
With color_cell\
.Interior.Color = vbGreen\
End With\
\
Case Is < 0\
With color_cell\
.Interior.Color = vbRed\
End With\
\
End Select\
\
Next g\
Next ws\
\
End Sub\
}