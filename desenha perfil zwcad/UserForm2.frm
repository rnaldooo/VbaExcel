VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "UserForm1"
   ClientHeight    =   1320
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8070
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'https://stackoverflow.com/questions/15535540/get-selected-value-of-a-combobox

Private Sub ComboBox3_Change()
    str_layer = ComboBox3.Value
End Sub

Private Sub CommandButton1_Click()
    cell.Offset(0, -1).Value = str_layer
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    Dim ii As Integer
    Dim pp As String
    Dim zcad As IZcadApplication
    Dim Thisdrawing As ZcadDocument
    ' ZWCAD**** - definindo desenho aberto
    Set zcad = GetObject(, Range("C2").Value)
    Dim aAD As ZWCAD.ZcadDocument: Set aAD = zcad.ActiveDocument:      aAD.Activate ' aAD - zcad document
    Dim i, il As Integer
    il = aAD.Layers.Count

    Dim l As Object
    Set l = CreateObject("System.Collections.ArrayList")

    For i = 0 To (il - 1)
        l.Add aAD.Layers(i).Name
        Worksheets("Layers").Cells(i + 1, 1).Value = aAD.Layers(i).Name
    Next i

    l.Sort


    Dim k As Variant
    For Each k In l
        ComboBox3.AddItem k
        ''# add the sorted items back to your combobox instead
        Debug.Print k
    Next k


    'For i = 0 To (il - 1)
    'ComboBox3.AddItem aAD.Layers(i).Name
    'Next i
    ComboBox3.ListIndex = 1

End Sub

Sub Chk_Item_SelectOrNot()
    MsgBox ComboBox1.Value
End Sub

