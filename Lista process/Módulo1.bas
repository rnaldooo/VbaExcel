Attribute VB_Name = "Módulo1"
Sub Test_AllRunningApps()
    Dim apps() As Variant
    apps() = AllRunningApps

    Range("A1").Resize(UBound(apps), 1).Value2 = WorksheetFunction.Transpose(apps)
    Range("A:A").Columns.AutoFit
End Sub

'Similar to: http://msdn.microsoft.com/en-us/library/aa393618%28VS.85%29.aspx
Public Function AllRunningApps() As Variant
    Dim strComputer As String
    Dim objServices As Object, objProcessSet As Object, Process As Object
    Dim oDic As Object, a() As Variant

    Set oDic = CreateObject("Scripting.Dictionary")

    strComputer = "."

    Set objServices = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
    Set objProcessSet = objServices.ExecQuery("SELECT Name, ProcessID FROM Win32_Process", , 48)

    For Each Process In objProcessSet
       If Not oDic.exists(Process.Name) Then oDic.Add Process.Name, Process.Name
    Next

    a() = oDic.keys

    Set objProcessSet = Nothing
    Set oDic = Nothing

    AllRunningApps = a()
End Function

