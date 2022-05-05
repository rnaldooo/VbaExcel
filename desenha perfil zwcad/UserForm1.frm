VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Excel --> Autocad(Layout)"
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6105
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub ComboBox1_Change()
    If ComboBox1.Value <> "" Then
        Excel.Worksheets(ComboBox1.Value).Select
    End If
End Sub

Private Sub ComboBox2_Change()
    Excel.Workbooks(ComboBox2.Value).Activate
    ComboBox1.Clear
    Dim ii As Integer
    Dim pp As String
    Label1.Caption = Excel.ActiveWorkbook.Name
    For ii = 1 To (Excel.ActiveWorkbook.Sheets.Count)
        ComboBox1.AddItem Excel.ActiveWorkbook.Sheets(ii).Name
    Next ii
    ComboBox1.ListIndex = (Excel.ActiveWorkbook.ActiveSheet.Index - 1)
    pp = Excel.Selection.Address
    TextBox1.Value = pp
End Sub

Private Sub ComboBox3_Change()

End Sub

Private Sub UserForm_Initialize()
    Dim ii As Integer
    Dim pp As String
    Label1.Caption = Excel.ActiveWorkbook.Name
    For ii = 1 To (Excel.ActiveWorkbook.Sheets.Count)
        ComboBox1.AddItem Excel.ActiveWorkbook.Sheets(ii).Name
    Next ii
    ComboBox1.ListIndex = (Excel.ActiveWorkbook.ActiveSheet.Index - 1)
    pp = Excel.Selection.Address
    TextBox1.Value = pp
    

    For ii = 1 To (Excel.Workbooks.Count)
        ComboBox2.AddItem Excel.Workbooks.Item(ii).Name
    Next ii
    ComboBox2.Value = (Excel.ActiveWorkbook.Name)
    
    
    Dim zcad As IZcadApplication
    Dim Thisdrawing As ZcadDocument
    ' ZWCAD**** - definindo desenho aberto
    Set zcad = GetObject(, Range("C2").Value)
    Dim aAD As ZWCAD.ZcadDocument: Set aAD = zcad.ActiveDocument:      aAD.Activate ' aAD - zcad document
    Dim i, il As Integer
    il = aAD.Layers.Count
    For i = 0 To (il - 1)
        ComboBox3.AddItem aAD.Layers(i).Name
    Next i
    ComboBox3.ListIndex = 1

End Sub

Private Sub CommandButton1_Click()
    TextBox1.Value = GetUserRange
End Sub

Private Sub CommandButton2_Click()
    Call null_cria_layer
End Sub

Private Sub CommandButton3_Click()
    Call null_cria_tabela
End Sub

