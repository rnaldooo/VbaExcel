Attribute VB_Name = "Módulo4"

Public str_layer As String
Public cell As Range

Sub formulario()
    Dim BR, BC As Long
    Dim LC As Integer
    'Dim cell As Range
    BR = ActiveSheet.Buttons(Application.Caller).TopLeftCell.Row
    BC = ActiveSheet.Buttons(Application.Caller).TopLeftCell.Column

    Set cell = Cells(BR, BC)

    UserForm2.Show
End Sub

Sub pegaponto()
    Dim BR, BC As Long
    Dim LC As Integer
    'Dim cell As Range
    BR = ActiveSheet.Buttons(Application.Caller).TopLeftCell.Row
    BC = ActiveSheet.Buttons(Application.Caller).TopLeftCell.Column

    Set cell = Cells(BR, BC)


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
    
    Worksheets("desenha perfil").Cells(BR - 2, BC).Value = Vponto(0)
    Worksheets("desenha perfil").Cells(BR - 1, BC).Value = Vponto(1)
    
mostrar:
    Excel.Visible = True


End Sub

' .net usaer referencia  mscorlib.dll

Sub test()
    Dim l As Object
  
  
    Set l = CreateObject("System.Collections.ArrayList")

    ''# these would be the items from your combobox, obviously
    ''# ... add them with a for loop
    l.Add "d"
    l.Add "c"
    l.Add "b"
    l.Add "a"

    l.Sort

    ''# now clear your combobox

    Dim k As Variant
    For Each k In l
        ''# add the sorted items back to your combobox instead
        Debug.Print k
    Next k

End Sub

Sub SearchListExample()
    'Define array list and variables
    Dim MyList As New ArrayList, Sp As Integer, Pos As Integer
    'Add new items including a duplicate
    MyList.Add "Item1"
    MyList.Add "Item2"
    MyList.Add "Item3"
    MyList.Add "Item1"
    'Test for “Item2” being in list - returns True
    MsgBox MyList.Contains("Item2")
    'Get index of non-existent value – returns -1
    MsgBox MyList.IndexOf("Item", 0)
    'Set the start position for the search to zero
    Sp = 0
    'Iterate through list to get all positions of ‘Item1”
    Do
        'Get the index position of the next ‘Item1’ based on the position in the variable ‘Sp’
        Pos = MyList.IndexOf("Item1", Sp)
        'If no further instances of ‘Item1’ are found then exit the loop
        If Pos = -1 Then Exit Do
        'Display the next instance found and the index position
        MsgBox MyList(Pos) & " at index " & Pos
        'Add 1 to the last found index value – this now becomes the new start position for the next search
        Sp = Pos + 1
    Loop
End Sub

Sub Ch4_ChangeHatchPatternSpace()
    Dim hatchObj As AcadHatch
    Dim patternName As String
    Dim PatternType As Long
    Dim bAssociativity As Boolean
    ' Define the hatch
    patternName = "ANSI31"
    PatternType = 0
    bAssociativity = True
    ' Create the associative Hatch object
    Set hatchObj = Thisdrawing.ModelSpace. _
                   AddHatch(PatternType, patternName, bAssociativity)
    ' Create the outer loop for the hatch.
    Dim outerLoop(0 To 0) As AcadEntity
    Dim center(0 To 2) As Double
    Dim radius As Double
    center(0) = 5
    center(1) = 3
    center(2) = 0
    radius = 3
    Set outerLoop(0) = Thisdrawing.ModelSpace. _
                       AddCircle(center, radius)
    hatchObj.AppendOuterLoop (outerLoop)
    hatchObj.Evaluate
    ' Change the spacing of the hatch pattern by
    ' adding 2 to the current spacing
    hatchObj.PatternSpace = hatchObj.PatternSpace + 2
    hatchObj.Evaluate
    Thisdrawing.Regen True
End Sub

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
'



Sub Example_AddHatch()
    ' This example creates an associative gradient hatch in model space.

    Dim hatchObj As AcadHatch
    Dim patternName As String
    Dim PatternType As Long
    Dim bAssociativity As Boolean

    ' Define the hatch
    patternName = "CYLINDER"
    PatternType = acPreDefinedGradient           '0
    bAssociativity = True

    ' Create the associative Hatch object in model space
    Set hatchObj = Thisdrawing.ModelSpace.AddHatch(PatternType, patternName, bAssociativity, acGradientObject)
    Dim col1 As AcadAcCmColor, col2 As AcadAcCmColor
    Set col1 = AcadApplication.GetInterfaceObject("AutoCAD.AcCmColor.16")
    Set col2 = AcadApplication.GetInterfaceObject("AutoCAD.AcCmColor.16")
    Call col1.SetRGB(255, 0, 0)
    Call col2.SetRGB(0, 255, 0)
    hatchObj.GradientColor1 = col1
    hatchObj.GradientColor2 = col2

    ' Create the outer boundary for the hatch (a circle)
    Dim outerLoop(0 To 0) As AcadEntity
    Dim center(0 To 2) As Double
    Dim radius As Double
    center(0) = 3: center(1) = 3: center(2) = 0
    radius = 1
    Set outerLoop(0) = Thisdrawing.ModelSpace.AddCircle(center, radius)

    ' Append the outerboundary to the hatch object, and display the hatch
    hatchObj.AppendOuterLoop (outerLoop)
    hatchObj.Evaluate
    Thisdrawing.Regen True
End Sub


