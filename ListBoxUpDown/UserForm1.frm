VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5865
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub MoveDown_Click()
Dim lCurrentListIndex As Long
Dim strRowSource As String
Dim strAddress As String
Dim strSheetName As String


With ListBox1
  If .ListIndex < 0 Or .ListIndex = .ListCount - 1 Then Exit Sub
    lCurrentListIndex = .ListIndex + 1
    strRowSource = .RowSource
    strAddress = Range(strRowSource).Address
    strSheetName = Range(strRowSource).Parent.Name
    .RowSource = vbNullString
        With Range(strRowSource)
            .Rows(lCurrentListIndex).Cut
            .Rows(lCurrentListIndex + 2).Insert Shift:=xlDown
        End With
     Sheets(strSheetName).Range(strAddress).Name = strRowSource
    .RowSource = strRowSource
    .Selected(lCurrentListIndex) = True
End With

End Sub

Private Sub MoveUp_Click()
Dim lCurrentListIndex As Long
Dim strRowSource As String
Dim strAddress As String
Dim strSheetName As String


With ListBox1
  If .ListIndex < 1 Then Exit Sub
    lCurrentListIndex = .ListIndex + 1
    strRowSource = .RowSource
    strAddress = Range(strRowSource).Address
    strSheetName = Range(strRowSource).Parent.Name
    .RowSource = vbNullString
        With Range(strRowSource)
            .Rows(lCurrentListIndex).Cut
            .Rows(lCurrentListIndex - 1).Insert Shift:=xlDown
        End With
     Sheets(strSheetName).Range(strAddress).Name = strRowSource
    .RowSource = strRowSource
    .Selected(lCurrentListIndex - 2) = True
End With

End Sub

