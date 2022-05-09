Sub copiablocoVparaH()
'
' Macro1 Macro
'
'
Dim Rorigem As Range
Set Rorigem = Range("e4")  'primeira célula
Dim Rultima As Range
Set Rultima = Range("j460")  'ultima primeira célula

Dim Rpassa As Range
Set Rpassa = Range("m4")  'primeira célula
Dim iquantasL As Integer
iquantasL = 43          'a cada quantas linhas
Dim iquantasC As Integer
iquantasC = 6          'quantas colunas

Dim iolinha As Integer
Dim iocoluna As Integer
iolinha = Rorigem.Row
iocoluna = Rorigem.Column

Dim iulinha As Integer
Dim iucoluna As Integer
iulinha = Rultima.Row
iucoluna = Rultima.Column

Dim iplinha As Integer
Dim ipcoluna As Integer
iplinha = Rpassa.Row
ipcoluna = Rpassa.Column

Dim il1, il2 As Integer
Dim ic1, ic2 As Integer

iquantasL = iquantasL - 1
iquantasC = iquantasC - 1
il1 = iolinha
ic1 = iocoluna
ic2 = ipcoluna
Dim iconta As Integer
iconta = 1
While il1 <= iulinha
Cells(iplinha - 1, ic2).Value = Cells(il1, 3).Value
Cells(iplinha - 1, ic2 + 1).Value = Cells(il1, 4).Value

Range(Cells(il1, iocoluna), Cells(il1 + iquantasL, iocoluna + iquantasC)).Select
Selection.Copy
il1 = il1 + (iquantasL + 1)
iconta = iconta + 1


Cells(iplinha, ic2).Select
ActiveSheet.Paste
Application.CutCopyMode = False



ic2 = ic2 + (iquantasC + 1)

 
 Wend
 
End Sub












