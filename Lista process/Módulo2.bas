Attribute VB_Name = "Módulo2"


Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Const MAX_LEN = 260
Public Function EnumWinProc(ByVal hwnd As Long, ByVal lParam As Long) As Long
Dim lRet As Long
Dim strBuffer As String

If IsWindowVisible(hwnd) Then
strBuffer = Space(MAX_LEN)
lRet = GetWindowText(hwnd, strBuffer, Len(strBuffer))
If lRet Then
UserForm1.ComboBox1.AddItem Left(strBuffer & " " & hwnd, lRet)
End If
End If

EnumWinProc = 1
End Function

'You can change the SQL to read Select Name, ProcessID FROM Win32_Process

'Then in your For Loop, to get the name use Process.Properties_("Name").value and Process.Properties_("ProcessID").value where needed.


