Attribute VB_Name = "Módulo2"
Sub DESENHA_CAD222()
    '--------------------------------------------------------------------------------------------------------------------------------
    ' reinaldo - 02/04/2020  - desenha polilinha no cad
    '--------------------------------------------------------------------------------------------------------------------------------
 
    Set Zcad1 = GetObject(, Worksheets("desenha perfil").Range("c2").Value)
    Dim aAD As ZWCAD.ZcadDocument
    Set aAD = ZWCAD.ActiveDocument
    aAD.Activate
    Dim Excel As Object
    Set Excel = GetObject(, "Excel.Application") ' define excel como o objeto
    AppActivate2 Zcad1.Caption
    
    On Error GoTo mostrar
    
    Dim Vponto As Variant
    Dim aADu As ZcadUtility
    Set aADu = aAD.Utility
    aADu.Prompt (Chr(10))
    aADu.Prompt ("SELECIONE O PONTO DE INSERÇÃO)")
    Vponto = aADu.GetPoint(, "selecione:")
    
    Dim pii As Double
    pii = 4 * Atn(1)
    
    Dim mx, my As Double
    mx = Worksheets("desenha perfil").Range("b17").Value
    my = Worksheets("desenha perfil").Range("C17").Value
    
    Dim curves(0 To 6) As ZcadEntity
    Dim aaa As ZcadEntity
    
    Dim arcObj0 As ZcadArc
    Dim plineObj1 As ZcadLWPolyline
    Dim arcObj2 As ZcadArc
    Dim L3 As ZcadLine
    Dim plineObj2 As ZcadLWPolyline
    Dim L5 As ZcadLine
    
    '0arco
    ' Define the arc
    Dim centerPoint0(0 To 2) As Double
    Dim radius0 As Double
    Dim startAngle0 As Double
    Dim endAngle0 As Double
    centerPoint0(0) = Worksheets("desenha perfil").Range("a22").Value * mx + Vponto(0)
    centerPoint0(1) = Worksheets("desenha perfil").Range("b22").Value * my + Vponto(1)
    centerPoint0(2) = 0#
    radius0 = Worksheets("desenha perfil").Range("d15").Value * mx '+ Vponto(0)
    startAngle0 = (pii * 2) / 4 * 3
    endAngle0 = pii * 2
    
    Set curves(0) = aAD.ModelSpace.AddArc(centerPoint0, radius0, startAngle0, endAngle0)
    Set arcObj0 = curves(0)
   
   
    '1fundo
    Dim pontos1() As Double
    Dim itam, ilinha, ipoli As Integer
    itam = 6
    ipoli = (itam * 2 - 1)
    ReDim pontos1(0 To ipoli)
    For ilinha = 1 To itam
        pontos1(ilinha * 2 - 2) = Worksheets("desenha perfil").Cells(ilinha + 24, 1).Value * mx + Vponto(0)
        pontos1(ilinha * 2 - 1) = Worksheets("desenha perfil").Cells(ilinha + 24, 2).Value * my + Vponto(1)
    Next ilinha
    Set curves(1) = aAD.ModelSpace.AddLightWeightPolyline(pontos1)
    Set plineObj1 = curves(1)
    aAD.Regen zcAllViewports
 
 
    '2arco
    ' Define the arc
    centerPoint0(0) = Worksheets("desenha perfil").Range("a33").Value * mx + Vponto(0)
    centerPoint0(1) = Worksheets("desenha perfil").Range("b33").Value * my + Vponto(1)
    centerPoint0(2) = 0#
    radius0 = Worksheets("desenha perfil").Range("d15").Value * mx '+ Vponto(0)
    startAngle0 = pii
    endAngle0 = (pii * 2) / 4 * 3
    Set curves(1) = aAD.ModelSpace.AddArc(centerPoint0, radius0, startAngle0, endAngle0)
    Set arcObj0 = curves(2)
 
    
    '3line
    Dim L3Point0(0 To 2) As Double
    Dim L3Point1(0 To 2) As Double
    L3Point0(0) = Worksheets("desenha perfil").Range("a36").Value * mx + Vponto(0)
    L3Point0(1) = Worksheets("desenha perfil").Range("b36").Value * my + Vponto(1)
    L3Point0(2) = 0#
    
    L3Point1(0) = Worksheets("desenha perfil").Range("a37").Value * mx + Vponto(0)
    L3Point1(1) = Worksheets("desenha perfil").Range("b37").Value * my + Vponto(1)
    L3Point1(2) = 0#
    
    Set curves(3) = aAD.ModelSpace.AddLine(L3Point0, L3Point1)
        
    '4arco
    ' Define the arc
    centerPoint0(0) = Worksheets("desenha perfil").Range("a40").Value * mx + Vponto(0)
    centerPoint0(1) = Worksheets("desenha perfil").Range("b40").Value * my + Vponto(1)
    centerPoint0(2) = 0#
    radius0 = Worksheets("desenha perfil").Range("d15").Value * mx '+ Vponto(0)
    startAngle0 = pii / 2
    endAngle0 = pii
    Set curves(4) = aAD.ModelSpace.AddArc(centerPoint0, radius0, startAngle0, endAngle0)
    Set arcObj0 = curves(4)
    
   
   
    '5cima
    Dim pontos5() As Double
    itam = 6
    ipoli = (itam * 2 - 1)
    ReDim pontos5(0 To ipoli)
    For ilinha = 1 To itam
        pontos5(ilinha * 2 - 2) = Worksheets("desenha perfil").Cells(ilinha + 42, 1).Value * mx + Vponto(0)
        pontos5(ilinha * 2 - 1) = Worksheets("desenha perfil").Cells(ilinha + 42, 2).Value * my + Vponto(1)
    Next ilinha
    Set curves(5) = aAD.ModelSpace.AddLightWeightPolyline(pontos5)
    Set plineObj1 = curves(5)
    aAD.Regen zcAllViewports
   
    '6arco
    ' Define the arc
    centerPoint0(0) = Worksheets("desenha perfil").Range("a51").Value * mx + Vponto(0)
    centerPoint0(1) = Worksheets("desenha perfil").Range("b51").Value * my + Vponto(1)
    centerPoint0(2) = 0#
    radius0 = Worksheets("desenha perfil").Range("d15").Value * mx '+ Vponto(0)
    startAngle0 = 0
    endAngle0 = pii / 2
    Set curves(6) = aAD.ModelSpace.AddArc(centerPoint0, radius0, startAngle0, endAngle0)
    Set arcObj0 = curves(6)
    aAD.Regen zcAllViewports
         
    Set aaa = aAD.ModelSpace.AddPolyline(curves)
    aaa.Closed = True
   

    Excel.Visible = True
    AppActivate Application.Caption
  
mostrar:
    Excel.Visible = True
  
End Sub

