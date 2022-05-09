Sub arredondar()
Dim guia As Worksheet
Dim selecao As Range
Set guia = ActiveSheet
Set selecao = Selection

For Each Cell In selecao
Cell.Value = Round(Cell.Value, 2)
Next

End Sub