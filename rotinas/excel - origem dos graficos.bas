Sub grafii()

Dim chart As ChartObject
For Each chart In ActiveSheet.ChartObjects
Dim ser As Series
For Each ser In chart.chart.SeriesCollection
Dim of As String
Dim nf As String
of = ser.Formula
nf = Replace(of, "[esfor?os vt1.xlsx]", "")
ser.Formula = nf
Next ser
Next chart

End Sub