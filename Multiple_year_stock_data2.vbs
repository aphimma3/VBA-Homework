Attribute VB_Name = "Module2"
Sub worksheet2015()
Dim ticker As String
Dim vol As Double
vol = 0

Range("I1").Value = "Ticker"
Range("J1").Value = "Total Stock Vol"

Dim summary_table_row As Integer
summary_table_row = 2

For i = 2 To 760192
If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
ticker = Cells(i, 1).Value
vol = vol + Cells(i, 7).Value

Range("I" & summary_table_row).Value = ticker
Range("J" & summary_table_row).Value = vol

summary_table_row = summary_table_row + 1

vol = 0

Else

vol = vol + Cells(i, 7).Value

End If

Next i
End Sub
