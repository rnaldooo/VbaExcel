Sub menosum()
Dim guia As Worksheet
Dim selecao As Range
Set guia = ActiveSheet
Set selecao = Selection

For Each Cell In selecao
Cell.Value = -Cell.Value
Next

End Sub