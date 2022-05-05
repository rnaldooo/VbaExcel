Attribute VB_Name = "Módulo1"
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
    Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
#Else
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
    Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
#End If

Private Zcad1 As ZcadApplication

Sub desenhaperfil()

    Set Zcad1 = GetObject(, Worksheets("desenha perfil").Range("c2").Value)
    Dim aAD As ZWCAD.ZcadDocument
    Set aAD = ZWCAD.ActiveDocument
    aAD.Activate
    Dim Excel As Object
    Set Excel = GetObject(, "Excel.Application") ' define excel como o objeto
    ' AppActivate Zcad1.Caption
    'AppActivate2 Zcad1.Caption
    
    On Error GoTo mostrar
    
    Dim Vponto, VPONTO2, VPONTOI As Variant
    Dim aADu As ZcadUtility
    Set aADu = aAD.Utility
    
    
    Dim VPONTOCTR As Variant
    VPONTOCTR = aAD.GetVariable("viewctr")
    
    
    
    aADu.Prompt (Chr(10))
    ' aADu.prompt ("clique em algum lugar")
    ' VPONTOI = aADu.GetPoint(VPONTO2, "selecione:")
    aADu.Prompt ("SELECIONE O PONTO DE INSERÇÃO)")
    Vponto = aADu.GetPoint(VPONTOCTR, "selecione:")
    
    Worksheets("desenha perfil").Range("L13").Value = Vponto(0)
    Worksheets("desenha perfil").Range("L14").Value = Vponto(1)
    
    
    
    Dim pii As Double
    pii = 4 * Atn(1)
    
    Dim mx, my As Double
    mx = Worksheets("desenha perfil").Range("b17").Value
    my = Worksheets("desenha perfil").Range("C17").Value
    
    Dim dro As Double
    dro = Worksheets("desenha perfil").Range("d4").Value
    
    Dim dbfi, dtfi, dbfs, dtfs, dh, dtw, draio, ddifbf    As Double
    Dim ilinn As Integer
    
    Dim llayer, lhachura, leixo As String
    'llayer = "0"
    llayer = Worksheets("desenha perfil").Range("K2").Value
    lhachura = Worksheets("desenha perfil").Range("K4").Value
    leixo = Worksheets("desenha perfil").Range("K5").Value
    
    Dim ccor As Integer
    ccor = 2
    
    ilinn = 8
    dh = Worksheets("desenha perfil").Range("D" & ilinn + 0 & "").Value * mx
    dtw = Worksheets("desenha perfil").Range("D" & ilinn + 1 & "").Value * mx
   
    dbfs = Worksheets("desenha perfil").Range("D" & ilinn + 2 & "").Value * mx
    dtfs = Worksheets("desenha perfil").Range("D" & ilinn + 3 & "").Value * mx
    
    dbfi = Worksheets("desenha perfil").Range("D" & ilinn + 4 & "").Value * mx
    dtfi = Worksheets("desenha perfil").Range("D" & ilinn + 5 & "").Value * mx
    
    
    draio = Worksheets("desenha perfil").Range("D" & ilinn + 7 & "").Value * mx
    ddifbf = (dbfs - dbfi) / 2
    ' centerPoint0(0) = Worksheets("desenha perfil").Range("a22").Value * mx + Vponto(0)
    
    
    'Dim curves(0 To 6) As ZcadEntity
    ' Dim aaa As ZcadEntity
    
    ' Dim arcObj0 As ZcadArc
    ' Dim plineObj1 As ZcadLWPolyline
    ' Dim arcObj2 As ZcadArc
    'Dim L3 As ZcadLine
    'Dim plineObj2 As ZcadLWPolyline
    ' Dim L5 As ZcadLine



    Dim VpontoIn(0 To 2) As Double
    Dim iin As Integer
    iin = Worksheets("desenha perfil").Range("f7").Value
      
    Select Case iin
    Case 1
        VpontoIn(0) = dbfi / 2: VpontoIn(1) = 0
    Case 2
        VpontoIn(0) = 0: VpontoIn(1) = 0
    Case 3
        VpontoIn(0) = -dbfi / 2: VpontoIn(1) = 0
    Case 4
        VpontoIn(0) = dbfi / 2: VpontoIn(1) = -dh
    Case 5
        VpontoIn(0) = 0: VpontoIn(1) = -dh
    Case 6
        VpontoIn(0) = -dbfi / 2: VpontoIn(1) = -dh
    Case 7
        VpontoIn(0) = 0: VpontoIn(1) = -dh / 2
    End Select
    
    
    Dim plineObj As ZcadLWPolyline
    Dim points(0 To 31) As Double

    ' Define the 2D polyline points
    points(0) = Vponto(0) + VpontoIn(0) - dbfi / 2:           points(1) = Vponto(1) + VpontoIn(1) + dtfi 'MESA   BAIXO   ESQUERDA    CHAPA
    points(2) = Vponto(0) + VpontoIn(0) - dbfi / 2:           points(3) = Vponto(1) + VpontoIn(1) 'MESA   BAIXO   ESQUERDA    FUNDO
            
    points(4) = Vponto(0) + VpontoIn(0) + dbfi / 2:           points(5) = Vponto(1) + VpontoIn(1) 'MESA   BAIXO   DIREITA FUNDO
    points(6) = Vponto(0) + VpontoIn(0) + dbfi / 2:           points(7) = Vponto(1) + VpontoIn(1) + dtfi 'MESA   BAIXO   DIREITA CHAPA
            
    points(8) = Vponto(0) + VpontoIn(0) + dtw / 2 + draio:    points(9) = Vponto(1) + VpontoIn(1) + dtfi 'ALMA   BAIXO   DIREITA CHAPA
    points(10) = Vponto(0) + VpontoIn(0) + dtw / 2:           points(11) = Vponto(1) + VpontoIn(1) + dtfi + draio 'ALMA   BAIXO   DIREITA ALMA
             
    points(12) = Vponto(0) + VpontoIn(0) + dtw / 2:           points(13) = Vponto(1) + VpontoIn(1) + dh - dtfs - draio 'ALMA   CIMA    DIREITA ALMA
    points(14) = Vponto(0) + VpontoIn(0) + dtw / 2 + draio:   points(15) = Vponto(1) + VpontoIn(1) + dh - dtfs 'ALMA   CIMA    DIREITA CHAPA
            
    points(16) = Vponto(0) + VpontoIn(0) + dbfi / 2 + ddifbf: points(17) = Vponto(1) + VpontoIn(1) + dh - dtfs 'MESA   BAIXO   DIREITA CHAPA
    points(18) = Vponto(0) + VpontoIn(0) + dbfi / 2 + ddifbf: points(19) = Vponto(1) + VpontoIn(1) + dh 'MESA   BAIXO   DIREITA FUNDO
            
    points(20) = Vponto(0) + VpontoIn(0) - dbfi / 2 - ddifbf: points(21) = Vponto(1) + VpontoIn(1) + dh 'MESA   BAIXO   ESQUERDA    FUNDO
    points(22) = Vponto(0) + VpontoIn(0) - dbfi / 2 - ddifbf: points(23) = Vponto(1) + VpontoIn(1) + dh - dtfs 'MESA   BAIXO   ESQUERDA    CHAPA
            
    points(24) = Vponto(0) + VpontoIn(0) - dtw / 2 - draio:   points(25) = Vponto(1) + VpontoIn(1) + dh - dtfs 'ALMA   CIMA    DIREITA CHAPA
    points(26) = Vponto(0) + VpontoIn(0) - dtw / 2:           points(27) = Vponto(1) + VpontoIn(1) + dh - dtfs - draio 'ALMA   CIMA    DIREITA ALMA
           
    points(28) = Vponto(0) + VpontoIn(0) - dtw / 2:           points(29) = Vponto(1) + VpontoIn(1) + dtfi + draio 'ALMA   BAIXO   DIREITA ALMA
    points(30) = Vponto(0) + VpontoIn(0) - dtw / 2 - draio:   points(31) = Vponto(1) + VpontoIn(1) + dtfi 'ALMA   BAIXO   DIREITA CHAPA
         
    ' Create a lightweight Polyline object in model space
    Set plineObj = aAD.ModelSpace.AddLightWeightPolyline(points)
    
    'set buldge to make it a semicircle
    Dim d_fator As Double
    d_fator = -0.41
    
    plineObj.SetBulge 4, d_fator
    plineObj.SetBulge 6, d_fator
    plineObj.SetBulge 12, d_fator
    plineObj.SetBulge 14, d_fator
    
    ' plineObj.SetBulge 1, 1
   
    'set the thickness
    ' plineObj.SetWidth 0, diam / 2, diam / 2
    ' plineObj.SetWidth 1, diam / 2, diam / 2
    
    plineObj.Closed = True
    plineObj.layer = llayer
    'plineObj.Color = ccor
    plineObj.Rotate Vponto, (dro * pi / 180)
    
    Dim outerLoop As Variant
    Dim outerLoopArray(0) As Object
    Set outerLoopArray(0) = plineObj
    outerLoop = outerLoopArray
    'MyHatch.AppendOuterLoop (outerLoop)
    
    Dim MyHatch As Object
    Set MyHatch = aAD.ModelSpace.AddHatch(zcHatchPatternTypePreDefined, "ANSI31", False)
    MyHatch.AppendOuterLoop (outerLoop)
    MyHatch.PatternScale = Worksheets("desenha perfil").Range("n4").Value '50#
    MyHatch.Evaluate
    MyHatch.layer = lhachura
    MyHatch.Update
    
   
   
         
    Dim L0 As ZcadLine
    Dim p11(0 To 2) As Double
    Dim p22(0 To 2) As Double
    
    'eixo
    If Worksheets("desenha perfil").Range("g7").Value Then
        p11(0) = Vponto(0) + VpontoIn(0):        p11(1) = Vponto(1) + VpontoIn(1):         p11(2) = 0#
        p22(0) = Vponto(0) + VpontoIn(0):       p22(1) = Vponto(1) + VpontoIn(1) + dh:      p11(2) = 0#
        Set L0 = aAD.ModelSpace.AddLine(p11, p22)
        L0.layer = leixo
        L0.Rotate Vponto, (dro * pi / 180)
        L0.Update
    End If
    


    
mostrar:
    Excel.Visible = True

End Sub

Sub desenhaenrijD()

    Set Zcad1 = GetObject(, Worksheets("desenha perfil").Range("c2").Value)
    Dim aAD As ZWCAD.ZcadDocument
    Set aAD = ZWCAD.ActiveDocument
    aAD.Activate
    Dim Excel As Object
    Set Excel = GetObject(, "Excel.Application") ' define excel como o objeto
    
    On Error GoTo mostrar
    
    Dim Vponto(0 To 2) As Double
    Dim VPONTO2, VPONTOI As Variant
    Dim aADu As ZcadUtility
    Set aADu = aAD.Utility
    
    aADu.Prompt (Chr(10))
    
    Vponto(0) = Worksheets("desenha perfil").Range("l13").Value
    Vponto(1) = Worksheets("desenha perfil").Range("l14").Value
    Vponto(2) = 0#
       
    Dim pii As Double
    pii = 4 * Atn(1)
    
    Dim mx, my As Double
    mx = Worksheets("desenha perfil").Range("b17").Value
    my = Worksheets("desenha perfil").Range("C17").Value
    
    Dim dro As Double
    dro = Worksheets("desenha perfil").Range("d4").Value
    
    Dim dbfi, dtfi, dbfs, dtfs, dh, dtw, draio, ddifbf    As Double
    Dim ilinn As Integer
    
    Dim llayer, lhachura, leixo, lchapa As String
    llayer = Worksheets("desenha perfil").Range("K2").Value
    lhachura = Worksheets("desenha perfil").Range("K4").Value
    leixo = Worksheets("desenha perfil").Range("K5").Value
    lchapa = Worksheets("desenha perfil").Range("K6").Value
    
    Dim ccor As Integer
    ccor = 2
    
    ilinn = 8
    dh = Worksheets("desenha perfil").Range("D" & ilinn + 0 & "").Value * mx
    dtw = Worksheets("desenha perfil").Range("D" & ilinn + 1 & "").Value * mx
   
    dbfs = Worksheets("desenha perfil").Range("D" & ilinn + 2 & "").Value * mx
    dtfs = Worksheets("desenha perfil").Range("D" & ilinn + 3 & "").Value * mx
    
    dbfi = Worksheets("desenha perfil").Range("D" & ilinn + 4 & "").Value * mx
    dtfi = Worksheets("desenha perfil").Range("D" & ilinn + 5 & "").Value * mx
    
    draio = Worksheets("desenha perfil").Range("D" & ilinn + 7 & "").Value * mx
    ddifbf = (dbfs - dbfi) / 2
    
    Dim dfolga, dchanf As Double
    dfolga = Worksheets("desenha perfil").Range("k11").Value * mx
    dchanf = Worksheets("desenha perfil").Range("k12").Value * mx
 
    Dim VpontoIn(0 To 2) As Double
    Dim iin As Integer
    iin = Worksheets("desenha perfil").Range("f7").Value
      
    Select Case iin
    Case 1
        VpontoIn(0) = dbfi / 2: VpontoIn(1) = 0
    Case 2
        VpontoIn(0) = 0: VpontoIn(1) = 0
    Case 3
        VpontoIn(0) = -dbfi / 2: VpontoIn(1) = 0
    Case 4
        VpontoIn(0) = dbfi / 2: VpontoIn(1) = -dh
    Case 5
        VpontoIn(0) = 0: VpontoIn(1) = -dh
    Case 6
        VpontoIn(0) = -dbfi / 2: VpontoIn(1) = -dh
    Case 7
        VpontoIn(0) = 0: VpontoIn(1) = -dh / 2
    End Select
    
    
    Dim plineObj As ZcadLWPolyline
    Dim points(0 To 11) As Double

    ' Define the 2D polyline points
    
    points(0) = Vponto(0) + VpontoIn(0) + dbfi / 2 + dfolga:            points(1) = Vponto(1) + VpontoIn(1) + dtfi 'MESA   BAIXO   DIREITA CHAPA
            
    points(2) = Vponto(0) + VpontoIn(0) + dtw / 2 + draio + dchanf:     points(3) = Vponto(1) + VpontoIn(1) + dtfi 'ALMA   BAIXO   DIREITA CHAPA
    points(4) = Vponto(0) + VpontoIn(0) + dtw / 2:                      points(5) = Vponto(1) + VpontoIn(1) + dtfi + draio + dchanf 'ALMA   BAIXO   DIREITA ALMA
             
    points(6) = Vponto(0) + VpontoIn(0) + dtw / 2:                      points(7) = Vponto(1) + VpontoIn(1) + dh - dtfs - draio - dchanf 'ALMA   CIMA    DIREITA ALMA
    points(8) = Vponto(0) + VpontoIn(0) + dtw / 2 + draio + dchanf:     points(9) = Vponto(1) + VpontoIn(1) + dh - dtfs 'ALMA   CIMA    DIREITA CHAPA
            
    points(10) = Vponto(0) + VpontoIn(0) + dbfi / 2 + ddifbf + dfolga: points(11) = Vponto(1) + VpontoIn(1) + dh - dtfs 'MESA   BAIXO   DIREITA CHAPA

    'points(22) = Vponto(0) + VpontoIn(0) - dbfi / 2 - ddifbf: points(23) = Vponto(1) + VpontoIn(1) + dh - dtfs 'MESA   BAIXO   ESQUERDA    CHAPA
            
    'points(24) = Vponto(0) + VpontoIn(0) - dtw / 2 - draio:   points(25) = Vponto(1) + VpontoIn(1) + dh - dtfs 'ALMA   CIMA    ESQUERDA CHAPA
    ' points(26) = Vponto(0) + VpontoIn(0) - dtw / 2:           points(27) = Vponto(1) + VpontoIn(1) + dh - dtfs - draio 'ALMA   CIMA    ESQUERDA ALMA
           
    ' points(28) = Vponto(0) + VpontoIn(0) - dtw / 2:           points(29) = Vponto(1) + VpontoIn(1) + dtfi + draio 'ALMA   BAIXO   ESQUERDA ALMA
    ' points(30) = Vponto(0) + VpontoIn(0) - dtw / 2 - draio:   points(31) = Vponto(1) + VpontoIn(1) + dtfi 'ALMA   BAIXO   ESQUERDA CHAPA
    ' points(0) = Vponto(0) + VpontoIn(0) - dbfi / 2:           points(1) = Vponto(1) + VpontoIn(1) + dtfi 'MESA   BAIXO   ESQUERDA    CHAPA

         
    ' Create a lightweight Polyline object in model space
    Set plineObj = aAD.ModelSpace.AddLightWeightPolyline(points)
    
    
    plineObj.Closed = True
    plineObj.layer = lchapa
    plineObj.Rotate Vponto, (dro * pi / 180)
    plineObj.Update


    Dim offsetObj As Variant
    Dim ssetname As String
    Dim sset As ZWCAD.ZcadSelectionSet
    ssetname = "s1"
    aAD.ActiveSelectionSet.Clear
    For Each sset In aAD.SelectionSets
        If sset.Name = ssetname Then
            sset.Delete
            Exit For
        End If
    Next sset
    Set offsetObj = aAD.SelectionSets.Add(ssetname)

    offsetObj = plineObj.Offset(dfolga)

    plineObj.Delete
    plineObj.Update
    
mostrar:
    Excel.Visible = True

End Sub

Sub desenhaenrijE()

    Set Zcad1 = GetObject(, Worksheets("desenha perfil").Range("c2").Value)
    Dim aAD As ZWCAD.ZcadDocument
    Set aAD = ZWCAD.ActiveDocument
    aAD.Activate
    Dim Excel As Object
    Set Excel = GetObject(, "Excel.Application") ' define excel como o objeto
    
    On Error GoTo mostrar
    
    Dim Vponto(0 To 2) As Double
    Dim VPONTO2, VPONTOI As Variant
    Dim aADu As ZcadUtility
    Set aADu = aAD.Utility
    
    aADu.Prompt (Chr(10))
    
    Vponto(0) = Worksheets("desenha perfil").Range("l13").Value
    Vponto(1) = Worksheets("desenha perfil").Range("l14").Value
    Vponto(2) = 0#
       
    Dim pii As Double
    pii = 4 * Atn(1)
    
    Dim mx, my As Double
    mx = Worksheets("desenha perfil").Range("b17").Value
    my = Worksheets("desenha perfil").Range("C17").Value
    
    Dim dro As Double
    dro = Worksheets("desenha perfil").Range("d4").Value
    
    Dim dbfi, dtfi, dbfs, dtfs, dh, dtw, draio, ddifbf    As Double
    Dim ilinn As Integer
    
    Dim llayer, lhachura, leixo, lchapa As String
    llayer = Worksheets("desenha perfil").Range("K2").Value
    lhachura = Worksheets("desenha perfil").Range("K4").Value
    leixo = Worksheets("desenha perfil").Range("K5").Value
    lchapa = Worksheets("desenha perfil").Range("K6").Value
    
    Dim ccor As Integer
    ccor = 2
    
    ilinn = 8
    dh = Worksheets("desenha perfil").Range("D" & ilinn + 0 & "").Value * mx
    dtw = Worksheets("desenha perfil").Range("D" & ilinn + 1 & "").Value * mx
   
    dbfs = Worksheets("desenha perfil").Range("D" & ilinn + 2 & "").Value * mx
    dtfs = Worksheets("desenha perfil").Range("D" & ilinn + 3 & "").Value * mx
    
    dbfi = Worksheets("desenha perfil").Range("D" & ilinn + 4 & "").Value * mx
    dtfi = Worksheets("desenha perfil").Range("D" & ilinn + 5 & "").Value * mx
    
    draio = Worksheets("desenha perfil").Range("D" & ilinn + 7 & "").Value * mx
    ddifbf = (dbfs - dbfi) / 2
    
    Dim dfolga, dchanf As Double
    dfolga = Worksheets("desenha perfil").Range("k11").Value * mx
    dchanf = Worksheets("desenha perfil").Range("k12").Value * mx
 
 
    Dim VpontoIn(0 To 2) As Double
    Dim iin As Integer
    iin = Worksheets("desenha perfil").Range("f7").Value
      
    Select Case iin
    Case 1
        VpontoIn(0) = dbfi / 2: VpontoIn(1) = 0
    Case 2
        VpontoIn(0) = 0: VpontoIn(1) = 0
    Case 3
        VpontoIn(0) = -dbfi / 2: VpontoIn(1) = 0
    Case 4
        VpontoIn(0) = dbfi / 2: VpontoIn(1) = -dh
    Case 5
        VpontoIn(0) = 0: VpontoIn(1) = -dh
    Case 6
        VpontoIn(0) = -dbfi / 2: VpontoIn(1) = -dh
    Case 7
        VpontoIn(0) = 0: VpontoIn(1) = -dh / 2
    End Select
    
    
    Dim plineObj As ZcadLWPolyline
    Dim points(0 To 11) As Double

    ' Define the 2D polyline points
    
    ' points(0) = Vponto(0) + VpontoIn(0) + dbfi / 2 + dfolga:          points(1) = Vponto(1) + VpontoIn(1) + dtfi 'MESA   BAIXO   DIREITA CHAPA
            
    ' points(2) = Vponto(0) + VpontoIn(0) + dtw / 2 + draio:    points(3) = Vponto(1) + VpontoIn(1) + dtfi 'ALMA   BAIXO   DIREITA CHAPA
    ' points(4) = Vponto(0) + VpontoIn(0) + dtw / 2:           points(5) = Vponto(1) + VpontoIn(1) + dtfi + draio 'ALMA   BAIXO   DIREITA ALMA
             
    ' points(6) = Vponto(0) + VpontoIn(0) + dtw / 2:           points(7) = Vponto(1) + VpontoIn(1) + dh - dtfs - draio 'ALMA   CIMA    DIREITA ALMA
    ' points(8) = Vponto(0) + VpontoIn(0) + dtw / 2 + draio:   points(9) = Vponto(1) + VpontoIn(1) + dh - dtfs 'ALMA   CIMA    DIREITA CHAPA
    '
    ' points(10) = Vponto(0) + VpontoIn(0) + dbfi / 2 + ddifbf + dfolga: points(11) = Vponto(1) + VpontoIn(1) + dh - dtfs 'MESA   BAIXO   DIREITA CHAPA


    points(0) = Vponto(0) + VpontoIn(0) - dbfi / 2 - ddifbf - dfolga:           points(1) = Vponto(1) + VpontoIn(1) + dh - dtfs 'MESA   BAIXO   ESQUERDA    CHAPA
    points(2) = Vponto(0) + VpontoIn(0) - dtw / 2 - draio - dchanf:             points(3) = Vponto(1) + VpontoIn(1) + dh - dtfs 'ALMA   CIMA    ESQUERDA CHAPA
    points(4) = Vponto(0) + VpontoIn(0) - dtw / 2:                              points(5) = Vponto(1) + VpontoIn(1) + dh - dtfs - draio - dchanf 'ALMA   CIMA    ESQUERDA ALMA
    points(6) = Vponto(0) + VpontoIn(0) - dtw / 2:                              points(7) = Vponto(1) + VpontoIn(1) + dtfi + draio + dchanf 'ALMA   BAIXO   ESQUERDA ALMA
    points(8) = Vponto(0) + VpontoIn(0) - dtw / 2 - draio - dchanf:             points(9) = Vponto(1) + VpontoIn(1) + dtfi 'ALMA   BAIXO   ESQUERDA CHAPA
    points(10) = Vponto(0) + VpontoIn(0) - dbfi / 2 - dfolga:                   points(11) = Vponto(1) + VpontoIn(1) + dtfi 'MESA   BAIXO   ESQUERDA    CHAPA

         
    ' Create a lightweight Polyline object in model space
    Set plineObj = aAD.ModelSpace.AddLightWeightPolyline(points)
    
    
    plineObj.Closed = True
    plineObj.layer = lchapa
    plineObj.Rotate Vponto, (dro * pi / 180)
    plineObj.Update


    Dim offsetObj As Variant
    Dim ssetname As String
    Dim sset As ZWCAD.ZcadSelectionSet
    ssetname = "s1"
    aAD.ActiveSelectionSet.Clear
    For Each sset In aAD.SelectionSets
        If sset.Name = ssetname Then
            sset.Delete
            Exit For
        End If
    Next sset
    Set offsetObj = aAD.SelectionSets.Add(ssetname)

    offsetObj = plineObj.Offset(dfolga)

    plineObj.Delete
    plineObj.Update
    
mostrar:
    Excel.Visible = True

End Sub

'-------------------------------------------------------------------------------------------
Function pi() As Variant                         ' - reinaldo - 25/03/2009 -
    pi = 4 * Atn(1)
End Function                                     '               dá o valor de pi

'-------------------------------------------------------------------------------------------



'      acHatch.SetDatabaseDefaults()
'      acHatch.SetHatchPattern(HatchPatternType.PreDefined, "ANSI31")
'      acHatch.Associative = True
'      acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl)
'      '' Evaluate the hatch
'      acHatch.EvaluateHatch (True)
'      '' Increase the pattern scale by 2 and re-evaluate the hatch
'      acHatch.PatternScale = acHatch.PatternScale + 2
'      acHatch.SetHatchPattern(acHatch.PatternType, acHatch.PatternName)
'      acHatch.EvaluateHatch (True)
'MyHatch.AppendInnerLoop (plineObj)



Sub vistamesa()

    Set Zcad1 = GetObject(, Worksheets("desenha perfil").Range("c2").Value)
    Dim aAD As ZWCAD.ZcadDocument
    Set aAD = ZWCAD.ActiveDocument
    aAD.Activate
    Dim Excel As Object
    Set Excel = GetObject(, "Excel.Application") ' define excel como o objeto
    '  AppActivate Zcad1.Caption
    
    On Error GoTo mostrar
    
    Dim Vponto, VPONTOI, VPONTO2, VPONTO3, VPONTO4  As Variant
    Dim aADu As ZcadUtility
    Set aADu = aAD.Utility
    
    Dim VPONTOCTR As Variant
    VPONTOCTR = aAD.GetVariable("viewctr")
    aADu.Prompt (Chr(10))
    ' aADu.prompt ("clique em algum lugar")
    ' VPONTOI = aADu.GetPoint(VPONTO2, "selecione:")
    
    aADu.Prompt ("SELECIONE O PONTO 1)")
    'PrikPnt = ActiveDocument.Utility.GetPoint(basePnt, PrikPointTExt)
    Vponto = aADu.GetPoint(VPONTOCTR, "selecione:")
    Worksheets("desenha perfil").Range("L13").Value = Vponto(0)
    Worksheets("desenha perfil").Range("L14").Value = Vponto(1)
    
    
    aADu.Prompt ("SELECIONE O PONTO 2)")
    VPONTO2 = aADu.GetPoint(Vponto, "selecione:")
    Worksheets("desenha perfil").Range("m13").Value = VPONTO2(0)
    Worksheets("desenha perfil").Range("m14").Value = VPONTO2(1)
    
    
    
    Dim pii As Double
    pii = 4 * Atn(1)
    
    Dim mx, my As Double
    mx = Worksheets("desenha perfil").Range("b17").Value
    my = Worksheets("desenha perfil").Range("C17").Value
    
    Dim dbfi, dtfi, dbfs, dtfs, dh, dtw, draio, ddifbf    As Double
    Dim ilinn As Integer
    
    Dim llayer, LL1, LL2, LL3 As String
    llayer = "0"
    LL1 = Worksheets("desenha perfil").Range("k2").Value
    LL2 = Worksheets("desenha perfil").Range("k3").Value
    LL3 = Worksheets("desenha perfil").Range("k5").Value
    
    Dim ccor As Integer
    ccor = 2
    
    ilinn = 8
    dh = Worksheets("desenha perfil").Range("D" & ilinn + 0 & "").Value * mx
    dtw = Worksheets("desenha perfil").Range("D" & ilinn + 1 & "").Value * mx
   
    dbfs = Worksheets("desenha perfil").Range("D" & ilinn + 2 & "").Value * mx
    dtfs = Worksheets("desenha perfil").Range("D" & ilinn + 3 & "").Value * mx
    
    dbfi = Worksheets("desenha perfil").Range("D" & ilinn + 4 & "").Value * mx
    dtfi = Worksheets("desenha perfil").Range("D" & ilinn + 5 & "").Value * mx
    
    
    draio = Worksheets("desenha perfil").Range("D" & ilinn + 7 & "").Value * mx
    ddifbf = (dbfs - dbfi) / 2
    ' centerPoint0(0) = Worksheets("desenha perfil").Range("a22").Value * mx + Vponto(0)
             
    Dim L0, L1, L2, L3, L4, L5 As ZcadLine
    Dim p11(0 To 2) As Double
    Dim p22(0 To 2) As Double
    
    'eixo
    If Worksheets("desenha perfil").Range("g7").Value Then
        p11(0) = Vponto(0):        p11(1) = Vponto(1):         p11(2) = 0#
        p22(0) = VPONTO2(0):       p22(1) = VPONTO2(1):        p11(2) = 0#
        Set L0 = aAD.ModelSpace.AddLine(p11, p22)
        L0.layer = LL3
        L0.Update
    End If
    
    'Dim points(0 To 3) As Double
    ' points(0) = Vponto(0):           points(1) = Vponto(1)
    ' points(2) = Vponto2(0):          points(3) = Vponto2(1)

    'Set L1 = zcLine(Vponto, Vponto2)
    'Set L1 = aAD.ModelSpace.AddLine(Vponto, Vponto2)
    
    Dim dx1, dx2, dy1, dy2, dang1, dang2, dang3, dangf, dmult As Double
       
    dy1 = Vponto(1) - VPONTO2(1)
    dx1 = Vponto(0) - VPONTO2(0)
    

    
    If dx1 = 0 Then
        dang1 = pii / 2
    
    Else
        dang1 = Atn(dy1 / dx1)
    End If
    
    Worksheets("desenha perfil").Range("L18").Value = dang1
        
    dang2 = dang1 + pii / 2
    dang3 = dang1 - pii / 2
        
    Dim isinal As Integer
    If (dx1 > 0 Or (dx1 = 0 And dy1 > 0)) Then
        isinal = -1
    Else
        isinal = 1
    End If
        
         
    Dim daf1(0 To 2), daf2(0 To 2), da1, da2 As Double
    da1 = Worksheets("desenha perfil").Range("c21").Value * mx * isinal
    da2 = -Worksheets("desenha perfil").Range("c22").Value * my * isinal
    daf1(0) = Vponto(0) + Cos(dang1) * da1:           daf1(1) = Vponto(1) + Sin(dang1) * da1:            daf1(2) = 0#
    daf2(0) = VPONTO2(0) + Cos(dang1) * da2:          daf2(1) = VPONTO2(1) + Sin(dang1) * da2:           daf2(2) = 0#
       
        
    'mesa 1
    dangf = dang2
    dmult = dbfs / 2#
    p11(0) = daf1(0) + Cos(dangf) * dmult:        p11(1) = daf1(1) + Sin(dangf) * dmult:         p11(2) = 0#
    p22(0) = daf2(0) + Cos(dangf) * dmult:        p22(1) = daf2(1) + Sin(dangf) * dmult:         p22(2) = 0#
  
    Set L1 = aAD.ModelSpace.AddLine(p11, p22)
    L1.layer = LL1
    L1.Update
    
    'mesa 2
    dangf = dang3
    dmult = dbfs / 2#
    p11(0) = daf1(0) + Cos(dangf) * dmult:        p11(1) = daf1(1) + Sin(dangf) * dmult:         p11(2) = 0#
    p22(0) = daf2(0) + Cos(dangf) * dmult:        p22(1) = daf2(1) + Sin(dangf) * dmult:         p22(2) = 0#
    Set L2 = aAD.ModelSpace.AddLine(p11, p22)
    L2.layer = LL1
    L2.Update

    'alma 1
    dangf = dang2
    dmult = dtw / 2#
    p11(0) = daf1(0) + Cos(dangf) * dmult:        p11(1) = daf1(1) + Sin(dangf) * dmult:         p11(2) = 0#
    p22(0) = daf2(0) + Cos(dangf) * dmult:        p22(1) = daf2(1) + Sin(dangf) * dmult:         p22(2) = 0#
    Set L3 = aAD.ModelSpace.AddLine(p11, p22)
    L3.layer = LL2
    L3.Update
    
    'alma 2
    dangf = dang3
    dmult = dtw / 2#
    p11(0) = daf1(0) + Cos(dangf) * dmult:        p11(1) = daf1(1) + Sin(dangf) * dmult:         p11(2) = 0#
    p22(0) = daf2(0) + Cos(dangf) * dmult:        p22(1) = daf2(1) + Sin(dangf) * dmult:         p22(2) = 0#
    Set L4 = aAD.ModelSpace.AddLine(p11, p22)
    L4.layer = LL2
    L4.Update
     

  
    
    
    ' Dim offsetObj As Variant
    ' Set offsetObj = aAD.SelectionSets.Add("test")
    ' offsetObj = L1.Offset(-dbfs)
    ' offsetObj = L1.Offset(-dbfs * 2)
    'offsetObj = L1.Offset(-dbfs * 5)
      
    'offrei ((1 + (Cos((((360 / nlados) * ia) * Pi) / 180))) / 2) * distancia1
     
    'L1 = L1.Offset(-dbfs)
    ' L2 = L1.Offset(-dbfs / 2#)
    
    'L3 = L1.Offset(-dbfs / 2#)
    'Set L2 = aAD.ModelSpace.AddLine()
    
    ' Set L3 = L1.Offset(-dbfs / 2#)
    ' Set L4 = L1.Offset(dtw / 2#)
    ' Set L5 = L1.Offset(-dtw / 2#)
     
    

    'Dim transMat(0 To 3, 0 To 3) As Double
    '  transMat(0, 0) = 0#: transMat(0, 1) = -1#: transMat(0, 2) = 0#: transMat(0, 3) = 0#
    '  transMat(1, 0) = 1#: transMat(1, 1) = 0#: transMat(1, 2) = 0#: transMat(1, 3) = 0#
    '  transMat(2, 0) = 0#: transMat(2, 1) = 0#: transMat(2, 2) = 1#: transMat(2, 3) = 0#
    '  transMat(3, 0) = 0#: transMat(3, 1) = 0#: transMat(3, 2) = 0#: transMat(3, 3) = 1#
    
    ' Transform the arc using the defined transformation matrix
    '  MsgBox "Transform the arc.", , "TransformBy Example"
    '  arcObj.TransformBy transMat

    
    
    ' plineObj.SetBulge 1, 1
   
    'set the thickness
    ' plineObj.SetWidth 0, diam / 2, diam / 2
    ' plineObj.SetWidth 1, diam / 2, diam / 2
    
    ' plineObj.layer = llayer
    ' plineObj.Color = ccor
    ' plineObj.Closed = True
    
mostrar:
    Excel.Visible = True

End Sub

Sub vistaalma()

    Set Zcad1 = GetObject(, Worksheets("desenha perfil").Range("c2").Value)
    Dim aAD As ZWCAD.ZcadDocument
    Set aAD = ZWCAD.ActiveDocument
    aAD.Activate
    Dim Excel As Object
    Set Excel = GetObject(, "Excel.Application") ' define excel como o objeto
    ' AppActivate Zcad1.Caption
       
    On Error GoTo mostrar
    
    'Both didn 't solve the problem unfortunately, same error still occurs when I use:
    '(ActiveDoc - 1 points to the right index)
    'Dim ActiveDoc1 As AcadDocument
    'Set ActiveDoc1 = Application.Documents.Item(ActiveDoc - 1)
    'PrikPnt = ActiveDoc1.Utility.GetPoint(Basepnt, PrikPointTexT)
    
    
    Dim Vponto, VPONTOI, VPONTO2, VPONTO3, VPONTO4 As Variant
    Dim aADu As ZcadUtility
    Set aADu = aAD.Utility
    
    Dim VPONTOCTR As Variant
    VPONTOCTR = aAD.GetVariable("viewctr")
    
    aADu.Prompt (Chr(10))
    
    ' aADu.prompt ("clique em algum lugar")
    ' VPONTOI = aADu.GetPoint(VPONTO2, "selecione:")
    aADu.Prompt ("SELECIONE O PONTO 1)")
    Vponto = aADu.GetPoint(VPONTOCTR, "selecione:")
    Worksheets("desenha perfil").Range("L13").Value = Vponto(0)
    Worksheets("desenha perfil").Range("L14").Value = Vponto(1)
    
    aADu.Prompt ("SELECIONE O PONTO 2)")
    VPONTO2 = aADu.GetPoint(Vponto, "selecione:")
    Worksheets("desenha perfil").Range("m13").Value = VPONTO2(0)
    Worksheets("desenha perfil").Range("m14").Value = VPONTO2(1)
    
    Dim pii As Double
    pii = 4 * Atn(1)
    
    Dim mx, my As Double
    mx = Worksheets("desenha perfil").Range("b17").Value
    my = Worksheets("desenha perfil").Range("C17").Value
    
    Dim dbfi, dtfi, dbfs, dtfs, dh, dtw, draio, ddifbf    As Double
    Dim ilinn As Integer
    
    Dim llayer, LL1, LL2, LL3 As String
    llayer = "0"
    LL1 = Worksheets("desenha perfil").Range("k2").Value
    LL2 = Worksheets("desenha perfil").Range("k3").Value
    LL3 = Worksheets("desenha perfil").Range("k5").Value
    
    Dim ccor As Integer
    ccor = 2
    
    ilinn = 8
    dh = Worksheets("desenha perfil").Range("D" & ilinn + 0 & "").Value * mx
    dtw = Worksheets("desenha perfil").Range("D" & ilinn + 1 & "").Value * mx
   
    dbfs = Worksheets("desenha perfil").Range("D" & ilinn + 2 & "").Value * mx
    dtfs = Worksheets("desenha perfil").Range("D" & ilinn + 3 & "").Value * mx
    
    dbfi = Worksheets("desenha perfil").Range("D" & ilinn + 4 & "").Value * mx
    dtfi = Worksheets("desenha perfil").Range("D" & ilinn + 5 & "").Value * mx
    
    
    draio = Worksheets("desenha perfil").Range("D" & ilinn + 7 & "").Value * mx
    ddifbf = (dbfs - dbfi) / 2
    ' centerPoint0(0) = Worksheets("desenha perfil").Range("a22").Value * mx + Vponto(0)
             
    Dim L0, L1, L2, L3, L4, L5 As ZcadLine
    Dim p11(0 To 2) As Double
    Dim p22(0 To 2) As Double
    
    'eixo
    If Worksheets("desenha perfil").Range("g7").Value Then
        p11(0) = Vponto(0):        p11(1) = Vponto(1):         p11(2) = 0#
        p22(0) = VPONTO2(0):       p22(1) = VPONTO2(1):        p11(2) = 0#
        Set L0 = aAD.ModelSpace.AddLine(p11, p22)
        L0.layer = LL3
        L0.Update
    End If
    
    'Dim points(0 To 3) As Double
    ' points(0) = Vponto(0):           points(1) = Vponto(1)
    ' points(2) = Vponto2(0):          points(3) = Vponto2(1)

    'Set L1 = zcLine(Vponto, Vponto2)
    'Set L1 = aAD.ModelSpace.AddLine(Vponto, Vponto2)
    
    Dim dx1, dx2, dy1, dy2, dang1, dang2, dang3, dangf, dmult As Double
       
    dy1 = Vponto(1) - VPONTO2(1)
    dx1 = Vponto(0) - VPONTO2(0)
    
    
    If dx1 = 0 Then
        dang1 = pii / 2
    
    Else
        dang1 = Atn(dy1 / dx1)
    End If
        
    Worksheets("desenha perfil").Range("L18").Value = dang1
        
    dang2 = dang1 + pii / 2
    dang3 = dang1 - pii / 2
        
    Dim isinal As Integer
    If (dx1 > 0 Or (dx1 = 0 And dy1 > 0)) Then
        isinal = -1
    Else
        isinal = 1
    End If
        
        
    Dim daf1(0 To 2), daf2(0 To 2), da1, da2 As Double
    da1 = Worksheets("desenha perfil").Range("c21").Value * mx * isinal
    da2 = -Worksheets("desenha perfil").Range("c22").Value * my * isinal
    daf1(0) = Vponto(0) + Cos(dang1) * da1:           daf1(1) = Vponto(1) + Sin(dang1) * da1:            daf1(2) = 0#
    daf2(0) = VPONTO2(0) + Cos(dang1) * da2:          daf2(1) = VPONTO2(1) + Sin(dang1) * da2:           daf2(2) = 0#
        
    'mesa 1
    dangf = dang2
    dmult = dh / 2#
    p11(0) = daf1(0) + Cos(dangf) * dmult:       p11(1) = daf1(1) + Sin(dangf) * dmult:        p11(2) = 0#
    p22(0) = daf2(0) + Cos(dangf) * dmult:       p22(1) = daf2(1) + Sin(dangf) * dmult:        p22(2) = 0#
  
    Set L1 = aAD.ModelSpace.AddLine(p11, p22)
    L1.layer = LL1
    L1.Update
    
    'mesa 2
    dangf = dang3
    dmult = dh / 2#
    p11(0) = daf1(0) + Cos(dangf) * dmult:       p11(1) = daf1(1) + Sin(dangf) * dmult:         p11(2) = 0#
    p22(0) = daf2(0) + Cos(dangf) * dmult:       p22(1) = daf2(1) + Sin(dangf) * dmult:         p22(2) = 0#
    Set L2 = aAD.ModelSpace.AddLine(p11, p22)
    L2.layer = LL1
    L2.Update

    'esp mesa 1
    dangf = dang2
    dmult = dh / 2# - dtfs
    p11(0) = daf1(0) + Cos(dangf) * (dmult):     p11(1) = daf1(1) + Sin(dangf) * dmult:         p11(2) = 0#
    p22(0) = daf2(0) + Cos(dangf) * (dmult):     p22(1) = daf2(1) + Sin(dangf) * dmult:         p22(2) = 0#
    Set L3 = aAD.ModelSpace.AddLine(p11, p22)
    L3.layer = LL1
    L3.Update
    
    'esp mesa 2
    dangf = dang3
    dmult = dh / 2# - dtfi
    p11(0) = daf1(0) + Cos(dangf) * dmult:       p11(1) = daf1(1) + Sin(dangf) * dmult:         p11(2) = 0#
    p22(0) = daf2(0) + Cos(dangf) * dmult:       p22(1) = daf2(1) + Sin(dangf) * dmult:         p22(2) = 0#
    Set L4 = aAD.ModelSpace.AddLine(p11, p22)
    L4.layer = LL1
    L4.Update
     

  
    
    
    ' Dim offsetObj As Variant
    ' Set offsetObj = aAD.SelectionSets.Add("test")
    ' offsetObj = L1.Offset(-dbfs)
    ' offsetObj = L1.Offset(-dbfs * 2)
    'offsetObj = L1.Offset(-dbfs * 5)
      
    'offrei ((1 + (Cos((((360 / nlados) * ia) * Pi) / 180))) / 2) * distancia1
     
    'L1 = L1.Offset(-dbfs)
    ' L2 = L1.Offset(-dbfs / 2#)
    
    'L3 = L1.Offset(-dbfs / 2#)
    'Set L2 = aAD.ModelSpace.AddLine()
    
    ' Set L3 = L1.Offset(-dbfs / 2#)
    ' Set L4 = L1.Offset(dtw / 2#)
    ' Set L5 = L1.Offset(-dtw / 2#)
     
    

    'Dim transMat(0 To 3, 0 To 3) As Double
    '  transMat(0, 0) = 0#: transMat(0, 1) = -1#: transMat(0, 2) = 0#: transMat(0, 3) = 0#
    '  transMat(1, 0) = 1#: transMat(1, 1) = 0#: transMat(1, 2) = 0#: transMat(1, 3) = 0#
    '  transMat(2, 0) = 0#: transMat(2, 1) = 0#: transMat(2, 2) = 1#: transMat(2, 3) = 0#
    '  transMat(3, 0) = 0#: transMat(3, 1) = 0#: transMat(3, 2) = 0#: transMat(3, 3) = 1#
    
    ' Transform the arc using the defined transformation matrix
    '  MsgBox "Transform the arc.", , "TransformBy Example"
    '  arcObj.TransformBy transMat

    
    
    ' plineObj.SetBulge 1, 1
   
    'set the thickness
    ' plineObj.SetWidth 0, diam / 2, diam / 2
    ' plineObj.SetWidth 1, diam / 2, diam / 2
    
    ' plineObj.layer = llayer
    ' plineObj.Color = ccor
    ' plineObj.Closed = True
    
mostrar:
    Excel.Visible = True

End Sub

Sub Example_OffsetPolyline()
    
    'Create the polyline
    Dim plineObj As ZcadLWPolyline
    Dim points(0 To 5) As Double
    points(0) = 1: points(1) = 1
    points(2) = 1: points(3) = 2
    points(4) = 2: points(5) = 2
    
    Set plineObj = Thisdrawing.ModelSpace.AddLightWeightPolyline(points)
    plineObj.Closed = True
    Thisdrawing.Regen zcActiveViewport
    Thisdrawing.Application.ZoomAll
     
    'Offset the polyline
    Dim offsetObj As Variant
    Set offsetObj = Thisdrawing.SelectionSets.Add("test")
    offsetObj = plineObj.Offset(0.25)
    
End Sub

Sub DESENHA_CAD()
    '--------------------------------------------------------------------------------------------------------------------------------
    ' reinaldo - 20/06/2011  - desenha polilinha no cad
    '--------------------------------------------------------------------------------------------------------------------------------
    ' Dim w As Window
    'Set w = Application.ActiveWindow
    'Dim i, ideci As Integer
    'Dim vtemp As Variant
    'ideci = Worksheets("desenha perfil").Range("j6").Value
    'ideci = 6
    'For i = 1 To ideci '  Worksheets("desenha perfil").Range("a22").Value
    ' vtemp = Worksheets("desenha perfil").Cells(22 + i, 1).Value 'WorksheetFunction.Round(Worksheets("desenha perfil").Cells(22 + i, 1).Value, ideci)
    ' Worksheets("desenha perfil").Cells(i, 2).Value = vtemp
    'Next i
    'Dim Thisdrawing As ZWcadDocument
    ' AUTOCAD**** - definindo desenho aberto
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
    '   aAD.Regen zcAllViewports
    '   aAD.ActiveViewport.Application.ZoomExtents
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
    Set plineObj1 = curves(1)                    'aAD.ModelSpace.AddLightWeightPolyline(pontos)
    aAD.Regen zcAllViewports
    '  aAD.ActiveViewport.Application.ZoomExtents
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
    ' aAD.Regen zcAllViewports
    ' aAD.ActiveViewport.Application.ZoomExtents
    
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
    '  aAD.Regen zcAllViewports
    '  aAD.ActiveViewport.Application.ZoomExtents
   
   
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
    Set plineObj1 = curves(5)                    'aAD.ModelSpace.AddLightWeightPolyline(pontos)
    aAD.Regen zcAllViewports
    '   aAD.ActiveViewport.Application.ZoomExtents
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
    '   aAD.ActiveViewport.Application.ZoomExtents
    
    
      
    Set aaa = aAD.ModelSpace.AddPolyline(curves)
    aaa.Closed = True
   

    Excel.Visible = True
    AppActivate Application.Caption
    '  aAD.Regen zcAllViewports
    '   aAD.ActiveViewport.Application.ZoomExtents
mostrar:
    Excel.Visible = True
    'Dim limpa As Boolean
    'limpa = Worksheets("desenha perfil").Range("q3").Value
    'If limpa Then
    '   Columns("A:B").Select    'LIMPAR
    '   Selection.ClearContents  'LIMPAR
    '   Range("A1").Select       'LIMPAR
    '  Else
    '  End If
    'AppActivate Excel.Caption, 0
    'Dim tit As String
    'tit = Application.Caption
    'AppActivate2 w.Caption
    ' AppActivate w.Caption
    'AppActivate tit
    'Excel.getfocus
End Sub

Sub AppActivate2(ByVal windowTitle As String)
    Dim hwnd As Long
    'first try to find the window
    hwnd = FindWindow(vbNullString, windowTitle)
    If hwnd > 0 Then
        'we found the window and now we have its handle. So we can bring it to foreground
        SetForegroundWindow hwnd
        'ShowWindow hwnd
    Else
        'window was not found. do whatever you want to do here in this case
    End If
End Sub

Sub Example_AddRegion()
    ' This example creates a region from an arc and a line
    Dim curves(0 To 1) As ZcadEntity
    Dim arcObj As ZcadArc
    ' Define the arc
    Dim centerPoint(0 To 2) As Double
    Dim radius As Double
    Dim startAngle As Double
    Dim endAngle As Double
    centerPoint(0) = 100#: centerPoint(1) = 100#: centerPoint(2) = 0#
    radius = 10#
    startAngle = 0
    endAngle = 3.141592
    Set curves(0) = Thisdrawing.ModelSpace.AddArc(centerPoint, radius, startAngle, endAngle)
    Set arcObj = curves(0)
    
    ' Define the line
    Dim p1 As Variant
    Dim p2 As Variant
    p1 = arcObj.StartPoint
    p2 = arcObj.EndPoint
    Set curves(1) = Thisdrawing.ModelSpace.AddLine(p1, p2)
        
    ' Create the region
    Dim regionObj As Variant
    regionObj = Thisdrawing.ModelSpace.AddRegion(curves)
    Thisdrawing.Regen zcActiveViewport
    Thisdrawing.Application.ZoomAll
    
End Sub

Sub mostrar()
    Dim Excel As Object
    Set Excel = GetObject(, "Excel.Application") ' define excel como o objeto
    Excel.Visible = True
End Sub

Sub TestLWPline()


    Set Zcad1 = GetObject(, Worksheets("desenha perfil").Range("c2").Value)
    Dim aAD As ZWCAD.ZcadDocument
    Set aAD = ZWCAD.ActiveDocument
    aAD.Activate
    Dim Excel As Object
    Set Excel = GetObject(, "Excel.Application") ' define excel como o objeto
    AppActivate2 Zcad1.Caption
    
    'On Error GoTo mostrar
    
    Dim Vponto As Variant
    Dim aADu As ZcadUtility
    Set aADu = aAD.Utility
    ' aADu.prompt (Chr(10))
    ' aADu.prompt ("SELECIONE O PONTO DE INSERÇÃO)")
    'Vponto = aADu.GetPoint(, "selecione:")


    'build the initial array for two points
    Dim Pt(0 To 3) As Double
    Pt(0) = 0#: Pt(1) = 0#: Pt(2) = 0#: Pt(3) = 10#
    'create a lwpline based on the two points
    Dim LWpline As Object
    Set LWpline = aAD.ModelSpace.AddLightWeightPolyline(Pt)
    'add further points
    Pt(0) = 10#: Pt(1) = 10#
    LWpline.AddVertex 2, Pt
    'add a the start point to lwpline again so
    'that the start and end points coincide but its Closed
    'property is false
    Pt(0) = 0#: Pt(1) = 0#
    LWpline.AddVertex 3, Pt
    'you can check here if it is really false
    Dim bClosed As Boolean
    bClosed = LWpline.Closed
    'add a hatch to modelspace
    Dim MyHatch As Object
    Set MyHatch = aAD.ModelSpace.AddHatch(zcHatchPatternTypePreDefined, "ANSI31", True)
    'create and append outer loop
    Dim outerLoop As Variant

    Dim outerLoopArray(0) As Object
    Set outerLoopArray(0) = LWpline

    outerLoop = outerLoopArray
    MyHatch.AppendOuterLoop (outerLoop)

End Sub

Sub notify(ByVal company As String, Optional ByVal office As String = "QJZ")
    If office = "QJZ" Then
        Debug.Print ("office not supplied -- using Headquarters")
        office = "Headquarters"
    End If
    ' Insert code to notify headquarters or specified office.
End Sub

Sub lreta1()

    Set Zcad1 = GetObject(, Worksheets("desenha perfil").Range("c2").Value)
    Dim aAD As ZWCAD.ZcadDocument
    Set aAD = ZWCAD.ActiveDocument
    aAD.Activate
    Dim Excel As Object
    Set Excel = GetObject(, "Excel.Application") ' define excel como o objeto
          
    On Error GoTo mostrar
        
    Dim Vponto(0 To 2), VPONTO2(0 To 2) As Double
   
    Dim aADu As ZcadUtility
    Set aADu = aAD.Utility
  
    aADu.Prompt (Chr(10))
    
    Vponto(0) = Worksheets("desenha perfil").Range("L13").Value
    Vponto(1) = Worksheets("desenha perfil").Range("L14").Value
    Vponto(2) = 0#
 
    VPONTO2(0) = Worksheets("desenha perfil").Range("m13").Value
    VPONTO2(1) = Worksheets("desenha perfil").Range("m14").Value
    VPONTO2(2) = 0#
    
    
    Dim pii As Double
    pii = 4 * Atn(1)
    
    Dim mx, my As Double
    mx = Worksheets("desenha perfil").Range("b17").Value
    my = Worksheets("desenha perfil").Range("C17").Value
    
    Dim dbfi, dtfi, dbfs, dtfs, dh, dtw, draio, ddifbf    As Double
    Dim ilinn As Integer
    
    Dim llayer, LL1 As String
    llayer = "0"
    LL1 = Worksheets("desenha perfil").Range("k7").Value

    
    ilinn = 8
    dh = Worksheets("desenha perfil").Range("D" & ilinn + 0 & "").Value * mx
    dtw = Worksheets("desenha perfil").Range("D" & ilinn + 1 & "").Value * mx
   
    dbfs = Worksheets("desenha perfil").Range("D" & ilinn + 2 & "").Value * mx
    dtfs = Worksheets("desenha perfil").Range("D" & ilinn + 3 & "").Value * mx
    
    dbfi = Worksheets("desenha perfil").Range("D" & ilinn + 4 & "").Value * mx
    dtfi = Worksheets("desenha perfil").Range("D" & ilinn + 5 & "").Value * mx
    
    
    draio = Worksheets("desenha perfil").Range("D" & ilinn + 7 & "").Value * mx
    ddifbf = (dbfs - dbfi) / 2
    ' centerPoint0(0) = Worksheets("desenha perfil").Range("a22").Value * mx + Vponto(0)
             
    Dim L1, L2, L3, L4 As ZcadLine
    Dim p11(0 To 2) As Double
    Dim p22(0 To 2) As Double

    
    Dim dx1, dx2, dy1, dy2, dang1, dang2, dang3, dangf, dmult As Double
       
    dy1 = Vponto(1) - VPONTO2(1)
    dx1 = Vponto(0) - VPONTO2(0)
    
    
    If dx1 = 0 Then
        dang1 = pii / 2
    
    Else
        dang1 = Atn(dy1 / dx1)
    End If
        
    dang2 = dang1 + pii / 2
    dang3 = dang1 - pii / 2
        
           
    Dim isinal As Integer
    If (dx1 > 0 Or (dx1 = 0 And dy1 > 0)) Then
        isinal = -1
    Else
        isinal = 1
    End If
        
        
    Dim daf1(0 To 2), daf2(0 To 2), da1, da2, da3 As Double
    
    
    da1 = Worksheets("desenha perfil").Range("c21").Value * mx * isinal
    da2 = -Worksheets("desenha perfil").Range("c22").Value * my * isinal
    
    da3 = Worksheets("desenha perfil").Range("c24").Value
    
    daf1(0) = Vponto(0) + Cos(dang1) * da1:           daf1(1) = Vponto(1) + Sin(dang1) * da1:            daf1(2) = 0#
    daf2(0) = VPONTO2(0) + Cos(dang1) * da2:          daf2(1) = VPONTO2(1) + Sin(dang1) * da2:           daf2(2) = 0#
        
    'mesa 1
    dangf = dang2
    dmult = dh / 2#
    
    p11(0) = daf1(0) + Cos(dangf) * (dmult + da3 * mx):    p11(1) = daf1(1) + Sin(dangf) * (dmult + da3 * my):       p11(2) = 0#
    p22(0) = daf1(0) + Cos(dangf) * (-dmult - da3 * mx):    p22(1) = daf1(1) + Sin(dangf) * (-dmult - da3 * my):       p22(2) = 0#
  
    'AddLightWeightPolyline
    Set L1 = aAD.ModelSpace.AddLine(p11, p22)
    L1.layer = LL1
    L1.Update
       
mostrar:
    Excel.Visible = True

End Sub

Sub lreta2()

    Set Zcad1 = GetObject(, Worksheets("desenha perfil").Range("c2").Value)
    Dim aAD As ZWCAD.ZcadDocument
    Set aAD = ZWCAD.ActiveDocument
    aAD.Activate
    Dim Excel As Object
    Set Excel = GetObject(, "Excel.Application") ' define excel como o objeto
          
    On Error GoTo mostrar
        
    Dim Vponto(0 To 2), VPONTO2(0 To 2) As Double
   
    Dim aADu As ZcadUtility
    Set aADu = aAD.Utility
  
    aADu.Prompt (Chr(10))
    
    Vponto(0) = Worksheets("desenha perfil").Range("L13").Value
    Vponto(1) = Worksheets("desenha perfil").Range("L14").Value
    Vponto(2) = 0#
 
    VPONTO2(0) = Worksheets("desenha perfil").Range("m13").Value
    VPONTO2(1) = Worksheets("desenha perfil").Range("m14").Value
    VPONTO2(2) = 0#
    
    
    Dim pii As Double
    pii = 4 * Atn(1)
    
    Dim mx, my As Double
    mx = Worksheets("desenha perfil").Range("b17").Value
    my = Worksheets("desenha perfil").Range("C17").Value
    
    Dim dbfi, dtfi, dbfs, dtfs, dh, dtw, draio, ddifbf    As Double
    Dim ilinn As Integer
    
    Dim llayer, LL1 As String
    llayer = "0"
    LL1 = Worksheets("desenha perfil").Range("k7").Value

    
    ilinn = 8
    dh = Worksheets("desenha perfil").Range("D" & ilinn + 0 & "").Value * mx
    dtw = Worksheets("desenha perfil").Range("D" & ilinn + 1 & "").Value * mx
   
    dbfs = Worksheets("desenha perfil").Range("D" & ilinn + 2 & "").Value * mx
    dtfs = Worksheets("desenha perfil").Range("D" & ilinn + 3 & "").Value * mx
    
    dbfi = Worksheets("desenha perfil").Range("D" & ilinn + 4 & "").Value * mx
    dtfi = Worksheets("desenha perfil").Range("D" & ilinn + 5 & "").Value * mx
    
    
    draio = Worksheets("desenha perfil").Range("D" & ilinn + 7 & "").Value * mx
    ddifbf = (dbfs - dbfi) / 2
    ' centerPoint0(0) = Worksheets("desenha perfil").Range("a22").Value * mx + Vponto(0)
             
    Dim L1, L2, L3, L4 As ZcadLine
    Dim p11(0 To 2) As Double
    Dim p22(0 To 2) As Double

    
    Dim dx1, dx2, dy1, dy2, dang1, dang2, dang3, dangf, dmult As Double
       
    dy1 = Vponto(1) - VPONTO2(1)
    dx1 = Vponto(0) - VPONTO2(0)
    
    
    If dx1 = 0 Then
        dang1 = pii / 2
    
    Else
        dang1 = Atn(dy1 / dx1)
    End If
        
    dang2 = dang1 + pii / 2
    dang3 = dang1 - pii / 2
        
        
           
    Dim isinal As Integer
    If (dx1 > 0 Or (dx1 = 0 And dy1 > 0)) Then
        isinal = -1
    Else
        isinal = 1
    End If
        
    Dim daf1(0 To 2), daf2(0 To 2), da1, da2, da3 As Double
    
    
    da1 = Worksheets("desenha perfil").Range("c21").Value * mx * isinal
    da2 = -Worksheets("desenha perfil").Range("c22").Value * my * isinal
    
    da3 = Worksheets("desenha perfil").Range("c24").Value
    
    daf1(0) = Vponto(0) + Cos(dang1) * da1:           daf1(1) = Vponto(1) + Sin(dang1) * da1:            daf1(2) = 0#
    daf2(0) = VPONTO2(0) + Cos(dang1) * da2:          daf2(1) = VPONTO2(1) + Sin(dang1) * da2:           daf2(2) = 0#
        
    'mesa 1
    dangf = dang2
    dmult = dh / 2#
    
    p11(0) = daf2(0) + Cos(dangf) * (dmult + da3 * mx):    p11(1) = daf2(1) + Sin(dangf) * (dmult + da3 * my):       p11(2) = 0#
    p22(0) = daf2(0) + Cos(dangf) * (-dmult - da3 * mx):    p22(1) = daf2(1) + Sin(dangf) * (-dmult - da3 * my):       p22(2) = 0#
  
    'AddLightWeightPolyline
    Set L1 = aAD.ModelSpace.AddLine(p11, p22)
    L1.layer = LL1
    L1.Update
       
mostrar:
    Excel.Visible = True

End Sub

Sub lcorte1()

    Set Zcad1 = GetObject(, Worksheets("desenha perfil").Range("c2").Value)
    Dim aAD As ZWCAD.ZcadDocument
    Set aAD = ZWCAD.ActiveDocument
    aAD.Activate
    Dim Excel As Object
    Set Excel = GetObject(, "Excel.Application") ' define excel como o objeto
          
    On Error GoTo mostrar
        
    Dim Vponto(0 To 2), VPONTO2(0 To 2) As Double
   
    Dim aADu As ZcadUtility
    Set aADu = aAD.Utility
  
    aADu.Prompt (Chr(10))
    
    Vponto(0) = Worksheets("desenha perfil").Range("L13").Value
    Vponto(1) = Worksheets("desenha perfil").Range("L14").Value
    Vponto(2) = 0#
 
    VPONTO2(0) = Worksheets("desenha perfil").Range("m13").Value
    VPONTO2(1) = Worksheets("desenha perfil").Range("m14").Value
    VPONTO2(2) = 0#
    
    
    Dim pii As Double
    pii = 4 * Atn(1)
    
    Dim mx, my As Double
    mx = Worksheets("desenha perfil").Range("b17").Value
    my = Worksheets("desenha perfil").Range("C17").Value
    
    Dim dbfi, dtfi, dbfs, dtfs, dh, dtw, draio, ddifbf    As Double
    Dim ilinn As Integer
    
    Dim llayer, LL1 As String
    llayer = "0"
    LL1 = Worksheets("desenha perfil").Range("k7").Value

    
    ilinn = 8
    dh = Worksheets("desenha perfil").Range("D" & ilinn + 0 & "").Value * mx
    dtw = Worksheets("desenha perfil").Range("D" & ilinn + 1 & "").Value * mx
   
    dbfs = Worksheets("desenha perfil").Range("D" & ilinn + 2 & "").Value * mx
    dtfs = Worksheets("desenha perfil").Range("D" & ilinn + 3 & "").Value * mx
    
    dbfi = Worksheets("desenha perfil").Range("D" & ilinn + 4 & "").Value * mx
    dtfi = Worksheets("desenha perfil").Range("D" & ilinn + 5 & "").Value * mx
    
    
    draio = Worksheets("desenha perfil").Range("D" & ilinn + 7 & "").Value * mx
    ddifbf = (dbfs - dbfi) / 2
    ' centerPoint0(0) = Worksheets("desenha perfil").Range("a22").Value * mx + Vponto(0)
             
    Dim L1, L2, L3, L4 As ZcadLine
    Dim p11(0 To 2) As Double
    Dim p22(0 To 2) As Double

    
    Dim dx1, dx2, dy1, dy2, dang1, dang2, dang3, dangf, dmult As Double
       
    dy1 = Vponto(1) - VPONTO2(1)
    dx1 = Vponto(0) - VPONTO2(0)
    
    
    If dx1 = 0 Then
        dang1 = pii / 2
    
    Else
        dang1 = Atn(dy1 / dx1)
    End If
        
    dang2 = dang1 + pii / 2
    dang3 = dang1 - pii / 2
        
           
    Dim isinal As Integer
    If (dx1 > 0 Or (dx1 = 0 And dy1 > 0)) Then
        isinal = -1
    Else
        isinal = 1
    End If
        
    Dim daf1(0 To 2), daf2(0 To 2), da1, da2, da3, da4 As Double
    
    
    da1 = Worksheets("desenha perfil").Range("c21").Value * mx * isinal
    da2 = -Worksheets("desenha perfil").Range("c22").Value * my * isinal
    
    da3 = Worksheets("desenha perfil").Range("c24").Value
    da4 = Worksheets("desenha perfil").Range("c25").Value
    
    daf1(0) = Vponto(0) + Cos(dang1) * da1:           daf1(1) = Vponto(1) + Sin(dang1) * da1:            daf1(2) = 0#
    daf2(0) = VPONTO2(0) + Cos(dang1) * da2:          daf2(1) = VPONTO2(1) + Sin(dang1) * da2:           daf2(2) = 0#
        
    'mesa 1
    dangf = dang2
    dmult = dh / 2#
  
  
    Dim plineObj As ZcadLWPolyline
    Dim points(0 To 11) As Double

    ' Define the 2D polyline points
    points(0) = daf1(0) + Cos(dangf) * (dmult + da3 * mx):                          points(1) = daf1(1) + Sin(dangf) * (dmult + da3 * my)
    points(2) = daf1(0) + Cos(dangf) * (da4 * mx):                                  points(3) = daf1(1) + Sin(dangf) * (da4 * my)
    points(4) = daf1(0) - Cos(dang1) * (da4 * mx):                                  points(5) = daf1(1) - Sin(dang1) * (da4 * mx)
    points(6) = daf1(0) + Cos(dang1) * (da4 * mx):                                  points(7) = daf1(1) + Sin(dang1) * (da4 * mx)
    points(8) = daf1(0) + Cos(dangf) * (-da4 * mx):                                 points(9) = daf1(1) + Sin(dangf) * (-da4 * my)
    points(10) = daf1(0) + Cos(dangf) * (-dmult - da3 * mx):                        points(11) = daf1(1) + Sin(dangf) * (-dmult - da3 * my)
         
    ' Create a lightweight Polyline object in model space
    Set plineObj = aAD.ModelSpace.AddLightWeightPolyline(points)
    
    plineObj.layer = LL1
    plineObj.Update
       
mostrar:
    Excel.Visible = True

End Sub

Sub lcorte2()

    Set Zcad1 = GetObject(, Worksheets("desenha perfil").Range("c2").Value)
    Dim aAD As ZWCAD.ZcadDocument
    Set aAD = ZWCAD.ActiveDocument
    aAD.Activate
    Dim Excel As Object
    Set Excel = GetObject(, "Excel.Application") ' define excel como o objeto
          
    On Error GoTo mostrar
        
    Dim Vponto(0 To 2), VPONTO2(0 To 2) As Double
   
    Dim aADu As ZcadUtility
    Set aADu = aAD.Utility
  
    aADu.Prompt (Chr(10))
    
    Vponto(0) = Worksheets("desenha perfil").Range("L13").Value
    Vponto(1) = Worksheets("desenha perfil").Range("L14").Value
    Vponto(2) = 0#
 
    VPONTO2(0) = Worksheets("desenha perfil").Range("m13").Value
    VPONTO2(1) = Worksheets("desenha perfil").Range("m14").Value
    VPONTO2(2) = 0#
    
    
    Dim pii As Double
    pii = 4 * Atn(1)
    
    Dim mx, my As Double
    mx = Worksheets("desenha perfil").Range("b17").Value
    my = Worksheets("desenha perfil").Range("C17").Value
    
    Dim dbfi, dtfi, dbfs, dtfs, dh, dtw, draio, ddifbf    As Double
    Dim ilinn As Integer
    
    Dim llayer, LL1 As String
    llayer = "0"
    LL1 = Worksheets("desenha perfil").Range("k7").Value

    
    ilinn = 8
    dh = Worksheets("desenha perfil").Range("D" & ilinn + 0 & "").Value * mx
    dtw = Worksheets("desenha perfil").Range("D" & ilinn + 1 & "").Value * mx
   
    dbfs = Worksheets("desenha perfil").Range("D" & ilinn + 2 & "").Value * mx
    dtfs = Worksheets("desenha perfil").Range("D" & ilinn + 3 & "").Value * mx
    
    dbfi = Worksheets("desenha perfil").Range("D" & ilinn + 4 & "").Value * mx
    dtfi = Worksheets("desenha perfil").Range("D" & ilinn + 5 & "").Value * mx
    
    
    draio = Worksheets("desenha perfil").Range("D" & ilinn + 7 & "").Value * mx
    ddifbf = (dbfs - dbfi) / 2
    ' centerPoint0(0) = Worksheets("desenha perfil").Range("a22").Value * mx + Vponto(0)
             
    Dim L1, L2, L3, L4 As ZcadLine
    Dim p11(0 To 2) As Double
    Dim p22(0 To 2) As Double

    
    Dim dx1, dx2, dy1, dy2, dang1, dang2, dang3, dangf, dmult As Double
       
    dy1 = Vponto(1) - VPONTO2(1)
    dx1 = Vponto(0) - VPONTO2(0)
    
    
    If dx1 = 0 Then
        dang1 = pii / 2
    
    Else
        dang1 = Atn(dy1 / dx1)
    End If
        
    dang2 = dang1 + pii / 2
    dang3 = dang1 - pii / 2
        
           
    Dim isinal As Integer
    If (dx1 > 0 Or (dx1 = 0 And dy1 > 0)) Then
        isinal = -1
    Else
        isinal = 1
    End If
        
    Dim daf1(0 To 2), daf2(0 To 2), da1, da2, da3, da4 As Double
    
    
    da1 = Worksheets("desenha perfil").Range("c21").Value * mx * isinal
    da2 = -Worksheets("desenha perfil").Range("c22").Value * my * isinal
    
    da3 = Worksheets("desenha perfil").Range("c24").Value
    da4 = Worksheets("desenha perfil").Range("c25").Value
    
    daf1(0) = Vponto(0) + Cos(dang1) * da1:           daf1(1) = Vponto(1) + Sin(dang1) * da1:            daf1(2) = 0#
    daf2(0) = VPONTO2(0) + Cos(dang1) * da2:          daf2(1) = VPONTO2(1) + Sin(dang1) * da2:           daf2(2) = 0#
        
    'mesa 1
    dangf = dang2
    dmult = dh / 2#
  
  
    Dim plineObj As ZcadLWPolyline
    Dim points(0 To 11) As Double

    ' Define the 2D polyline points
    points(0) = daf2(0) + Cos(dangf) * (dmult + da3 * mx):                          points(1) = daf2(1) + Sin(dangf) * (dmult + da3 * my)
    points(2) = daf2(0) + Cos(dangf) * (da4 * mx):                                  points(3) = daf2(1) + Sin(dangf) * (da4 * my)
    points(4) = daf2(0) - Cos(dang1) * (da4 * mx):                                  points(5) = daf2(1) - Sin(dang1) * (da4 * mx)
    points(6) = daf2(0) + Cos(dang1) * (da4 * mx):                                  points(7) = daf2(1) + Sin(dang1) * (da4 * mx)
    points(8) = daf2(0) + Cos(dangf) * (-da4 * mx):                                 points(9) = daf2(1) + Sin(dangf) * (-da4 * my)
    points(10) = daf2(0) + Cos(dangf) * (-dmult - da3 * mx):                        points(11) = daf2(1) + Sin(dangf) * (-dmult - da3 * my)
         
    ' Create a lightweight Polyline object in model space
    Set plineObj = aAD.ModelSpace.AddLightWeightPolyline(points)
    
    plineObj.layer = LL1
    plineObj.Update
       
mostrar:
    Excel.Visible = True

End Sub

Sub doispontos()

    Set Zcad1 = GetObject(, Worksheets("desenha perfil").Range("c2").Value)
    Dim aAD As ZWCAD.ZcadDocument
    Set aAD = ZWCAD.ActiveDocument
    aAD.Activate
    Dim Excel As Object
    Set Excel = GetObject(, "Excel.Application") ' define excel como o objeto
    ' AppActivate Zcad1.Caption
       
    On Error GoTo mostrar
    
    'Both didn 't solve the problem unfortunately, same error still occurs when I use:
    '(ActiveDoc - 1 points to the right index)
    'Dim ActiveDoc1 As AcadDocument
    'Set ActiveDoc1 = Application.Documents.Item(ActiveDoc - 1)
    'PrikPnt = ActiveDoc1.Utility.GetPoint(Basepnt, PrikPointTexT)
    
    
    Dim Vponto, VPONTOI, VPONTO2, VPONTO3, VPONTO4 As Variant
    Dim aADu As ZcadUtility
    Set aADu = aAD.Utility
    
    Dim VPONTOCTR As Variant
    VPONTOCTR = aAD.GetVariable("viewctr")
    
    aADu.Prompt (Chr(10))
    
    ' aADu.prompt ("clique em algum lugar")
    ' VPONTOI = aADu.GetPoint(VPONTO2, "selecione:")
    aADu.Prompt ("SELECIONE O PONTO 1)")
    Vponto = aADu.GetPoint(VPONTOCTR, "selecione:")
    Worksheets("desenha perfil").Range("L13").Value = Vponto(0)
    Worksheets("desenha perfil").Range("L14").Value = Vponto(1)
    
    aADu.Prompt ("SELECIONE O PONTO 2)")
    VPONTO2 = aADu.GetPoint(Vponto, "selecione:")
    Worksheets("desenha perfil").Range("m13").Value = VPONTO2(0)
    Worksheets("desenha perfil").Range("m14").Value = VPONTO2(1)
    
mostrar:
    Excel.Visible = True
End Sub

Sub lreta()

    Set Zcad1 = GetObject(, Worksheets("desenha perfil").Range("c2").Value)
    Dim aAD As ZWCAD.ZcadDocument
    Set aAD = ZWCAD.ActiveDocument
    aAD.Activate
    Dim Excel As Object
    Set Excel = GetObject(, "Excel.Application") ' define excel como o objeto
          
    On Error GoTo mostrar
        
    Dim Vponto, VPONTOI, VPONTO2, VPONTO3, VPONTO4 As Variant
    Dim aADu As ZcadUtility
    Set aADu = aAD.Utility
    
    Dim VPONTOCTR As Variant
    VPONTOCTR = aAD.GetVariable("viewctr")
        
    aADu.Prompt (Chr(10))
    
    aADu.Prompt ("SELECIONE O PONTO 1)")
    Vponto = aADu.GetPoint(VPONTOCTR, "selecione:")
    Worksheets("desenha perfil").Range("L13").Value = Vponto(0)
    Worksheets("desenha perfil").Range("L14").Value = Vponto(1)
    
    aADu.Prompt ("SELECIONE O PONTO 2)")
    VPONTO2 = aADu.GetPoint(Vponto, "selecione:")
    Worksheets("desenha perfil").Range("m13").Value = VPONTO2(0)
    Worksheets("desenha perfil").Range("m14").Value = VPONTO2(1)
    
    
    Dim pii As Double
    pii = 4 * Atn(1)
    
    Dim mx, my As Double
    mx = Worksheets("desenha perfil").Range("b17").Value
    my = Worksheets("desenha perfil").Range("C17").Value
    
    Dim dbfi, dtfi, dbfs, dtfs, dh, dtw, draio, ddifbf    As Double
    Dim ilinn As Integer
    
    Dim llayer, LL1 As String
    llayer = "0"
    LL1 = Worksheets("desenha perfil").Range("k7").Value

    
    ilinn = 8
    dh = Worksheets("desenha perfil").Range("D" & ilinn + 0 & "").Value * mx
    dtw = Worksheets("desenha perfil").Range("D" & ilinn + 1 & "").Value * mx
   
    dbfs = Worksheets("desenha perfil").Range("D" & ilinn + 2 & "").Value * mx
    dtfs = Worksheets("desenha perfil").Range("D" & ilinn + 3 & "").Value * mx
    
    dbfi = Worksheets("desenha perfil").Range("D" & ilinn + 4 & "").Value * mx
    dtfi = Worksheets("desenha perfil").Range("D" & ilinn + 5 & "").Value * mx
    
    
    draio = Worksheets("desenha perfil").Range("D" & ilinn + 7 & "").Value * mx
    ddifbf = (dbfs - dbfi) / 2
    ' centerPoint0(0) = Worksheets("desenha perfil").Range("a22").Value * mx + Vponto(0)
             
    Dim L1, L2, L3, L4 As ZcadLine
    Dim p11(0 To 2) As Double
    Dim p22(0 To 2) As Double

    
    Dim dx1, dx2, dy1, dy2, dang1, dang2, dang3, dangf, dmult As Double
       
    dy1 = Vponto(1) - VPONTO2(1)
    dx1 = Vponto(0) - VPONTO2(0)
    
    
    If dx1 = 0 Then
        dang1 = pii / 2
    
    Else
        dang1 = Atn(dy1 / dx1)
    End If
        
    dang2 = dang1 + pii / 2
    dang3 = dang1 - pii / 2
        
           
    Dim isinal As Integer
    If (dx1 > 0 Or (dx1 = 0 And dy1 > 0)) Then
        isinal = -1
    Else
        isinal = 1
    End If
        
        
    Dim daf1(0 To 2), daf2(0 To 2), da1, da2, da3 As Double
    
    
    da1 = 0#
    da2 = 0#
    
    da3 = -Worksheets("desenha perfil").Range("c24").Value * isinal
    
    daf1(0) = Vponto(0):            daf1(1) = Vponto(1):             daf1(2) = 0#
    daf2(0) = VPONTO2(0):           daf2(1) = VPONTO2(1):            daf2(2) = 0#
        
    'mesa 1
    dangf = dang1
    dmult = 0#
    
    p11(0) = daf1(0) + Cos(dangf) * (dmult + da3 * mx):    p11(1) = daf1(1) + Sin(dangf) * (dmult + da3 * my):       p11(2) = 0#
    p22(0) = daf2(0) + Cos(dangf) * (-dmult - da3 * mx):    p22(1) = daf2(1) + Sin(dangf) * (-dmult - da3 * my):       p22(2) = 0#
  
    'AddLightWeightPolyline
    Set L1 = aAD.ModelSpace.AddLine(p11, p22)
    L1.layer = LL1
    L1.Update
       
mostrar:
    Excel.Visible = True

End Sub

Sub lcorte()

    Set Zcad1 = GetObject(, Worksheets("desenha perfil").Range("c2").Value)
    Dim aAD As ZWCAD.ZcadDocument
    Set aAD = ZWCAD.ActiveDocument
    aAD.Activate
    Dim Excel As Object
    Set Excel = GetObject(, "Excel.Application") ' define excel como o objeto
          
    On Error GoTo mostrar
        
    Dim Vponto, VPONTOI, VPONTO2, VPONTO3, VPONTO4 As Variant
    Dim aADu As ZcadUtility
    Set aADu = aAD.Utility
    
    Dim VPONTOCTR As Variant
    VPONTOCTR = aAD.GetVariable("viewctr")
        
    aADu.Prompt (Chr(10))
    
    aADu.Prompt ("SELECIONE O PONTO 1)")
    Vponto = aADu.GetPoint(VPONTOCTR, "selecione:")
    Worksheets("desenha perfil").Range("L13").Value = Vponto(0)
    Worksheets("desenha perfil").Range("L14").Value = Vponto(1)
    
    aADu.Prompt ("SELECIONE O PONTO 2)")
    VPONTO2 = aADu.GetPoint(Vponto, "selecione:")
    Worksheets("desenha perfil").Range("m13").Value = VPONTO2(0)
    Worksheets("desenha perfil").Range("m14").Value = VPONTO2(1)
    
    
    Dim pii As Double
    pii = 4 * Atn(1)
    
    Dim mx, my As Double
    mx = Worksheets("desenha perfil").Range("b17").Value
    my = Worksheets("desenha perfil").Range("C17").Value
    
    Dim dbfi, dtfi, dbfs, dtfs, dh, dtw, draio, ddifbf    As Double
    Dim ilinn As Integer
    
    Dim llayer, LL1 As String
    llayer = "0"
    LL1 = Worksheets("desenha perfil").Range("k7").Value

    
    ilinn = 8
    dh = Worksheets("desenha perfil").Range("D" & ilinn + 0 & "").Value * mx
    dtw = Worksheets("desenha perfil").Range("D" & ilinn + 1 & "").Value * mx
   
    dbfs = Worksheets("desenha perfil").Range("D" & ilinn + 2 & "").Value * mx
    dtfs = Worksheets("desenha perfil").Range("D" & ilinn + 3 & "").Value * mx
    
    dbfi = Worksheets("desenha perfil").Range("D" & ilinn + 4 & "").Value * mx
    dtfi = Worksheets("desenha perfil").Range("D" & ilinn + 5 & "").Value * mx
    
    
    draio = Worksheets("desenha perfil").Range("D" & ilinn + 7 & "").Value * mx
    ddifbf = (dbfs - dbfi) / 2
    ' centerPoint0(0) = Worksheets("desenha perfil").Range("a22").Value * mx + Vponto(0)
             
    Dim L1, L2, L3, L4 As ZcadLine
    Dim p11(0 To 2) As Double
    Dim p22(0 To 2) As Double

    
    Dim dx1, dx2, dy1, dy2, dang1, dang2, dang3, dangf, dmult As Double
       
    dy1 = Vponto(1) - VPONTO2(1)
    dx1 = Vponto(0) - VPONTO2(0)
    
    
    If dx1 = 0 Then
        dang1 = pii / 2
    
    Else
        dang1 = Atn(dy1 / dx1)
    End If
        
    dang2 = dang1 + pii / 2
    dang3 = dang1 - pii / 2
        
           
    Dim isinal As Integer
    If (dx1 > 0 Or (dx1 = 0 And dy1 > 0)) Then
        isinal = -1
    Else
        isinal = 1
    End If
        
        
    Dim daf1(0 To 2), daf2(0 To 2), dafm(0 To 2), da1, da2, da3, da4 As Double
    
    
    da1 = 0#
    da2 = 0#
    
    da3 = -Worksheets("desenha perfil").Range("c24").Value * isinal
    da4 = -Worksheets("desenha perfil").Range("c25").Value * isinal
    
    daf1(0) = Vponto(0):                daf1(1) = Vponto(1):                daf1(2) = 0#
    daf2(0) = VPONTO2(0):               daf2(1) = VPONTO2(1):               daf2(2) = 0#
    dafm(0) = (daf1(0) + daf2(0)) / 2:  dafm(1) = (daf1(1) + daf2(1)) / 2:  dafm(2) = (daf1(2) + daf2(2)) / 2
        
    'mesa 1
    dangf = dang2
    dmult = 0#
    
    
    Dim plineObj As ZcadLWPolyline
    Dim points(0 To 11) As Double

    ' Define the 2D polyline points
    points(0) = daf1(0) + Cos(dang1) * (dmult + da3 * mx):                          points(1) = daf1(1) + Sin(dang1) * (dmult + da3 * my)
    points(2) = dafm(0) + Cos(dang1) * (da4 * mx):                                  points(3) = dafm(1) + Sin(dang1) * (da4 * my)
    points(4) = dafm(0) - Cos(dangf) * (da4 * mx):                                  points(5) = dafm(1) - Sin(dangf) * (da4 * mx)
    points(6) = dafm(0) + Cos(dangf) * (da4 * mx):                                  points(7) = dafm(1) + Sin(dangf) * (da4 * mx)
    points(8) = dafm(0) + Cos(dang1) * (-da4 * mx):                                 points(9) = dafm(1) + Sin(dang1) * (-da4 * my)
    points(10) = daf2(0) + Cos(dang1) * (-dmult - da3 * mx):                        points(11) = daf2(1) + Sin(dang1) * (-dmult - da3 * my)
         
    ' Create a lightweight Polyline object in model space
    Set plineObj = aAD.ModelSpace.AddLightWeightPolyline(points)
    
    plineObj.layer = LL1
    plineObj.Update
    
    
    
  
    'AddLightWeightPolyline
    Set L1 = aAD.ModelSpace.AddLine(p11, p22)
    L1.layer = LL1
    L1.Update
       
mostrar:
    Excel.Visible = True

End Sub

