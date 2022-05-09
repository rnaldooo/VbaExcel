' Note you need to add the reference Microsoft Scripting Runtime in Tools >> References. Change the Range("A1:A5") t
'A | 1
'X | 1
'C | 2
'D | 2
'E | 2
'
'
Option Explicit
Option Base 1

Dim Data As Dictionary

Sub GetCombinations()

    Dim dataObj As Variant
    Dim returnData As Variant
    Set Data = New Dictionary
    Dim i As Double

    dataObj = Range("A1:B5").Value2

    ' Group Data
    For i = 1 To UBound(dataObj) Step 1

        If (Data.Exists(dataObj(i, 2))) Then
            Data(dataObj(i, 2)) = Data(dataObj(i, 2)) & "|" & dataObj(i, 1)
        Else
            Data.Add dataObj(i, 2), dataObj(i, 1)
        End If

    Next i

    ' Extract combinations from groups
    returnData = CalculateCombinations().Keys()

    Range("G1").Resize(UBound(returnData) + 1, 1) = Application.WorksheetFunction.Transpose(returnData)

End Sub

Private Function CalculateCombinations() As Dictionary

    Dim i As Double, j As Double
    Dim datum As Variant, pieceInner As Variant, pieceOuter As Variant
    Dim Combo As New Dictionary
    Dim splitData() As String

    For Each datum In Data.Items

        splitData = Split(datum, "|")
        For Each pieceOuter In splitData
            For Each pieceInner In splitData

                If (pieceOuter <> pieceInner) Then

                    If (Not Combo.Exists(pieceOuter & "|" & pieceInner)) Then
                        Combo.Add pieceOuter & "|" & pieceInner, vbNullString
                    End If

                End If

            Next pieceInner
        Next pieceOuter

    Next datum

    Set CalculateCombinations = Combo

End Function
