Sub ShowInstalledFonts()
    Const StartRow As Integer = 4
    Dim cbcFontName As CommandBarControl, cbrFontCmd As CommandBar, strFormula As String
    Dim strFontName As String, i As Long, lngFontCount As Long, intFontSize As Integer
 
    intFontSize = 10
 
    If intFontSize = 0 Then Exit Sub
    If intFontSize < 8 Then intFontSize = 8
    If intFontSize > 30 Then intFontSize = 30
 
    Set cbcFontName = Application.CommandBars("Formatting").FindControl(ID:=1728)
    'Create a temp CommandBar if Font control is missing
    If cbcFontName Is Nothing Then
        Set cbrFontCmd = Application.CommandBars.Add("TempFontNamesCtrl", _
            msoBarFloating, False, True)
        Set cbcFontName = cbrFontCmd.Controls.Add(ID:=1728)
    End If
    Application.ScreenUpdating = False
    lngFontCount = cbcFontName.ListCount
    Workbooks.Add
    ' Column A - font names
    ' Column B - font example
    For i = 0 To cbcFontName.ListCount - 1
        strFontName = cbcFontName.List(i + 1)
        Application.StatusBar = "Listing font " & _
            Format(i / (lngFontCount - 1), "0 %") & " " & _
            strFontName & "..."
        Cells(i + StartRow, 1).Formula = strFontName
        With Cells(i + StartRow, 2)
            strFormula = "abcdefghijklmnopqrstuvwxyz"
            If Application.International(xlCountrySetting) = 47 Then
                strFormula = strFormula & "זרו"
            End If
            strFormula = strFormula & UCase(strFormula)
            strFormula = strFormula & "1234567890"
            .Formula = strFormula
            .Font.Name = strFontName
        End With
    Next i
    Application.StatusBar = False
    If Not cbrFontCmd Is Nothing Then cbrFontCmd.Delete
    Set cbrFontCmd = Nothing
    Set cbcFontName = Nothing
    ' Column heading
    Columns(1).AutoFit
    With Range("A1")
        .Formula = "Installed fonts:"
        .Font.Bold = True
        .Font.Size = 14
    End With
    With Range("A3")
        .Formula = "Font Name:"
        .Font.Bold = True
        .Font.Size = 12
    End With
    With Range("B3")
        .Formula = "Font Example:"
        .Font.Bold = True
        .Font.Size = 12
    End With
    With Range("B" & StartRow & ":B" & _
        StartRow + lngFontCount)
        .Font.Size = intFontSize
    End With
    With Range("A" & StartRow & ":B" & _
        StartRow + lngFontCount)
        .VerticalAlignment = xlVAlignCenter
    End With
    Range("A4").Select
    ActiveWindow.FreezePanes = True
    Range("A2").Select
    ActiveWorkbook.Saved = True
End Sub

