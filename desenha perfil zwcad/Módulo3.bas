Attribute VB_Name = "Módulo3"
Sub DESENHA_CAD333()
    '--------------------------------------------------------------------------------------------------------------------------------
    ' reinaldo - 02/04/2020  - desenha polilinha no cad
    '--------------------------------------------------------------------------------------------------------------------------------
 
    Set Zcad1 = GetObject(, Worksheets("desenha perfil").Range("c2").Value)
    Dim aAD As ZWCAD.ZcadDocument
    Set aAD = ZWCAD.ActiveDocument:      aAD.Activate
    Dim Excel As Object:      Set Excel = GetObject(, "Excel.Application") ' define excel como o objeto
    AppActivate2 Zcad1.Caption
    
    On Error GoTo mostrar
    
    Dim Vponto As Variant
    Dim aADu As ZcadUtility:     Set aADu = aAD.Utility
    aADu.Prompt (Chr(10))
    aADu.Prompt ("SELECIONE O PONTO DE INSERÇÃO)")
    Vponto = aADu.GetPoint(, "selecione:")       'ponto de insersao
    
    Dim pii As Double:      pii = 4 * Atn(1)
    
    Dim mx As Double:       mx = Worksheets("desenha perfil").Range("b17").Value ' escala x
    Dim my As Double:       my = Worksheets("desenha perfil").Range("C17").Value ' escala y
    
    
    Dim arcObj0, arcObj2 As ZcadArc
    Dim plineObj, plineObj1, plineObj2 As ZcadLWPolyline
    Dim L3, L5 As ZcadLine
    Dim points(0 To 7) As Double
    Dim iponto As Integer: iponto = 0
    
    '0arco
    ' Define the arc

  

    points(iponto) = Worksheets("desenha perfil").Range("a22").Value * mx + Vponto(0):      iponto = iponto + 1 '0
    points(iponto) = Worksheets("desenha perfil").Range("b22").Value * my + Vponto(1):      iponto = iponto + 1 '1
    
    'radius0 = Worksheets("desenha perfil").Range("d15").Value * mx '+ Vponto(0)
    ' startAngle0 = (pii * 2) / 4 * 3
    ' endAngle0 = pii * 2
  
   
    '1fundo
    Dim itam, ilinha, ipoli, iatu As Integer
    ilinha = 1
    For iatu = iponto To (iponto + 6)
        points(iponto) = Worksheets("desenha perfil").Cells(ilinha + 24, 1).Value * mx + Vponto(0):      iponto = iponto + 1
        points(iponto) = Worksheets("desenha perfil").Cells(ilinha + 24, 2).Value * my + Vponto(1):      iponto = iponto + 1
        ilinha = ilinha + 1
    Next iatu
    
       
    '2arco
    ' Define the arc
    points(iponto) = Worksheets("desenha perfil").Range("a33").Value * mx + Vponto(0): iponto = iponto + 1 '0
    points(iponto) = Worksheets("desenha perfil").Range("b33").Value * my + Vponto(1): iponto = iponto + 1
    
    '  radius0 = Worksheets("desenha perfil").Range("d15").Value * mx '+ Vponto(0)
    ' startAngle0 = pii
    '  endAngle0 = (pii * 2) / 4 * 3
   
     
    '3line

    points(iponto) = Worksheets("desenha perfil").Range("a36").Value * mx + Vponto(0): iponto = iponto + 1
    points(iponto) = Worksheets("desenha perfil").Range("b36").Value * my + Vponto(1): iponto = iponto + 1
    
    
    points(iponto) = Worksheets("desenha perfil").Range("a37").Value * mx + Vponto(0): iponto = iponto + 1
    points(iponto) = Worksheets("desenha perfil").Range("b37").Value * my + Vponto(1): iponto = iponto + 1
    
    
   
        
    '4arco
    ' Define the arc
    points(iponto) = Worksheets("desenha perfil").Range("a40").Value * mx + Vponto(0): iponto = iponto + 1
    points(iponto) = Worksheets("desenha perfil").Range("b40").Value * my + Vponto(1): iponto = iponto + 1
    
    'radius0 = Worksheets("desenha perfil").Range("d15").Value * mx '+ Vponto(0)
    ' startAngle0 = pii / 2
    'endAngle0 = pii
    
    
   
   
    '5cima
 
    ilinha = 1
    For iatu = iponto To (iponto + 6)
        pontos5(ilinha * 2 - 2) = Worksheets("desenha perfil").Cells(ilinha + 42, 1).Value * mx + Vponto(0): iponto = iponto + 1
        pontos5(ilinha * 2 - 1) = Worksheets("desenha perfil").Cells(ilinha + 42, 2).Value * my + Vponto(1): iponto = iponto + 1
        ilinha = ilinha + 1
    Next iatu
  
   
    '6arco
    ' Define the arc
    centerPoint0(0) = Worksheets("desenha perfil").Range("a51").Value * mx + Vponto(0): iponto = iponto + 1
    centerPoint0(1) = Worksheets("desenha perfil").Range("b51").Value * my + Vponto(1): iponto = iponto + 1
   
    'radius0 = Worksheets("desenha perfil").Range("d15").Value * mx '+ Vponto(0)
    ' startAngle0 = 0
    ' endAngle0 = pii / 2
   
    
    ' Define the 2D polyline points
    points(0) = ponto0(0): points(1) = ponto0(1)
    points(2) = ponto0(2): points(3) = ponto0(1)
    points(4) = ponto0(2): points(5) = ponto0(3)
    points(6) = ponto0(0): points(7) = ponto0(3)

    ' Create a lightweight Polyline object in model space
    Set plineObj = aAD.ModelSpace.AddLightWeightPolyline(points)
          
    '  plineObj.Closed = True
    'set the thickness
       
    ' plineObj.Layer = llayer
    ' plineObj.Color = ccor
    
    
    aAD.Regen zcAllViewports
         
    Set aaa = aAD.ModelSpace.AddPolyline(curves)
    aaa.Closed = True
   

    Excel.Visible = True
    AppActivate Application.Caption
  
mostrar:
    Excel.Visible = True
  
End Sub

