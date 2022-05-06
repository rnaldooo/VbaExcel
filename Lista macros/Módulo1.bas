Attribute VB_Name = "Módulo1"
Option Explicit


'Dim VBAEditor As VBIDE.VBE
'Dim VBProj As VBIDE.VBProject
'Dim VBComp As VBIDE.VBComponent
'Dim CodeMod As VBIDE.CodeModule

'Set VBAEditor = Application.VBE
'''''''''''''''''''''''''''''''''''''''''''
'Set VBProj = VBAEditor.ActiveVBProject
' or
'Set VBProj = Application.Workbooks("Book1.xls").VBProject
'''''''''''''''''''''''''''''''''''''''''''
'Set VBComp = ActiveWorkbook.VBProject.VBComponents("Module1")
' or
'Set VBComp = VBProj.VBComponents("Module1")
'''''''''''''''''''''''''''''''''''''''''''
'Set CodeMod = ActiveWorkbook.VBProject.VBComponents("Module1").CodeModule
' or
'Set CodeMod = VBComp.CodeModule



Sub ListProcedures()



        Dim VBProj As VBIDE.VBProject
        Dim VBComp As VBIDE.VBComponent
        Dim CodeMod As VBIDE.CodeModule
        Dim LineNum As Long
        Dim NumLines As Long
        Dim WS As Worksheet
        Dim Rng As Range
        Dim ProcName As String
        Dim ProcKind As VBIDE.vbext_ProcKind
        
        Set VBProj = ActiveWorkbook.VBProject
        Set VBComp = VBProj.VBComponents("Module1")
        Set CodeMod = VBComp.CodeModule
        
        Set WS = ActiveWorkbook.Worksheets("Sheet1")
        Set Rng = WS.Range("A1")
        
        With CodeMod
            LineNum = .CountOfDeclarationLines + 1
            ProcName = .ProcOfLine(LineNum, ProcKind)
            Do Until LineNum >= .CountOfLines
                Rng(1, 1).Value = ProcName
                Rng(1, 2).Value = ProcKindString(ProcKind)
                
                Set Rng = Rng(2, 1)
                LineNum = LineNum + .ProcCountLines(ProcName, ProcKind) + 1
                
                ProcName = .ProcOfLine(LineNum, ProcKind)
            Loop
        End With
    End Sub
    
    
    Function ProcKindString(ProcKind As VBIDE.vbext_ProcKind) As String
        Select Case ProcKind
            Case vbext_pk_Get
                ProcKindString = "Property Get"
            Case vbext_pk_Let
                ProcKindString = "Property Let"
            Case vbext_pk_Set
                ProcKindString = "Property Set"
            Case vbext_pk_Proc
                ProcKindString = "Sub Or Function"
            Case Else
                ProcKindString = "Unknown Type: " & CStr(ProcKind)
        End Select
    End Function





Dim mAppXL As Excel.Application

Dim mWb As Workbook

 '

Sub ListCodeNames()

Dim sFolder As String, bSubfolders As Boolean

Dim i As Long

Dim strStatusMsg As String

On Error GoTo Error_Handler

With Application.FileDialog(msoFileDialogFolderPicker)

   .InitialView = msoFileDialogViewList

    If .Show = -1 Then

        sFolder = .SelectedItems(1)

    Else

        MsgBox "Process terminated by user", vbCritical, "Terminated"

        Exit Sub

    End If

End With

bSubfolders = MsgBox("Search also in subfolders?", vbQuestion + vbYesNo) = vbYes

With Application.FileSearch

    .NewSearch

    .LookIn = sFolder

    .SearchSubFolders = bSubfolders

    .Filename = "*.xls"

    If .Execute() > 0 Then

        Set mAppXL = New Excel.Application

        Set mWb = Workbooks.Add

        mWb.Sheets(1).[A1].Resize(, 3) = Array("File", "Module", "Procedure")

        For i = 1 To .FoundFiles.Count

           strStatusMsg = "Processing " & i & " of " & .FoundFiles.Count & ": " & .FoundFiles(i)

           Application.StatusBar = strStatusMsg

           ListProcedures .FoundFiles(i)

        Next i

    Else

       MsgBox "No files found.", vbCritical

    End If

End With

ExitHere:

    'Cleanup

    Application.StatusBar = False

    mWb.Sheets(1).Columns("A:C").AutoFit

    mAppXL.Quit: Set mAppXL = Nothing

    MsgBox "Done.", vbInformation

Exit Sub

Error_Handler:

    MsgBox "Error " & Err.Number & ":" & vbLf & Err.Description, vbCritical, "Error"

    Resume ExitHere

End Sub



Sub ListProceduresggg(FilePath As String)

    Dim wb As Workbook

    Dim VBComp As VBComponent

    Dim VBCodeMod As CodeModule

    Dim StartLine As Long

    Dim ProcName As String

    '

    Set wb = mAppXL.Workbooks.Open(FilePath)

    For Each VBComp In wb.VBProject.VBComponents

        Set VBCodeMod = VBComp.CodeModule

        '

        With VBCodeMod

            StartLine = .CountOfDeclarationLines + 1

            Do Until StartLine >= .CountOfLines

            mWb.Sheets(1).[a65536].End(3).Offset(1).Resize(, 3) = _
                Array(FilePath, VBComp.Name, .ProcOfLine(StartLine, vbext_pk_Proc))
                StartLine = StartLine + _
                  .ProcCountLines(.ProcOfLine(StartLine, _
                   vbext_pk_Proc), vbext_pk_Proc)

            Loop

        End With

    '

    Next VBComp

    wb.Close False

End Sub



Sub FolderList()
'
' Example Macro to list the files contained in a folder.
'

Dim x, fs
Dim i As Integer
Dim y As Integer
Dim Folder, MyName, TotalFiles, Response, PrintResponse, AgainResponse

On Error Resume Next

Folder:
' Prompt the user for the folder to list.
x = InputBox("What folder do you want to list?" & Chr$(13) & Chr$(13) _
   & "For example: C:\My Documents")

If x = "" Or x = " " Then
    Response = MsgBox("Either you did not type a folder name correctly" _
    & Chr$(13) & "or you clicked Cancel. Do you want to quit?" _
    & Chr$(13) & Chr$(13) & _
    "If you want to type a folder name, click No." & Chr$(13) & _
    "If you want to quit, click Yes.", vbYesNo)

        If Response = "6" Then
            End
        Else
            GoTo Folder
        End If
Else

' Test if folder exists.
Set Folder = CreateObject("Scripting.filesystemobject")
If Folder.folderexists(x) = "True" Then

' Search the specified folder for files and type the listing in the ' document.
    
    With Application.FileSearch
        Set fs = Application.FileSearch
        fs.NewSearch
            With fs.PropertyTests
                .Add Name:="Files of Type", _
                Condition:=msoConditionFileTypeAllFiles, _
                Connector:=msoConnectorOr
            End With
        .LookIn = x
        .Execute
        TotalFiles = .FoundFiles.Count
        
                If TotalFiles <> 0 Then
                
                    ' Create a new document for the file listing.
                    Application.Documents.Add
                    ActiveDocument.ActiveWindow.View = wdPrintView
                    
                    ' Set tabs.
                    Selection.WholeStory
                    Selection.ParagraphFormat.TabStops.ClearAll
                    ActiveDocument.DefaultTabStop = InchesToPoints(0.5)
                    Selection.ParagraphFormat.TabStops.Add _
                       Position:=InchesToPoints(3), _
                       Alignment:=wdAlignTabLeft, _
                       Leader:=wdTabLeaderSpaces
                    Selection.ParagraphFormat.TabStops.Add _
                       Position:=InchesToPoints(4), _
                       Alignment:=wdAlignTabLeft, _
                       Leader:=wdTabLeaderSpaces
                    
                    ' Type the file list headings.
                    Selection.TypeText "File Listing of the "
                    
                    With Selection.Font
                        .Allcaps = True
                        .Bold = True
                    End With
                    Selection.TypeText x
                    With Selection.Font
                        .Allcaps = False
                        .Bold = False
                    End With
                    
                    Selection.TypeText " folder!" & Chr$(13)
                    
                    With Selection.Font
                        .Underline = wdUnderlineSingle
                    End With
                    With Selection
                        .TypeText Chr$(13)
                        .TypeText "File Name" & vbTab & "File Size" _
                           & vbTab & "File Date/Time" & Chr$(13)
                        .TypeText Chr$(13)
                    End With
                        With Selection.Font
                        .Underline = wdUnderlineNone
                    End With
        
                Else
                    MsgBox ("There are no files in the folder!" & _
                       "Please type another folder to list.")
                    GoTo Folder
                End If
        
        For i = 1 To TotalFiles
            MyName = .FoundFiles.Item(i)
            .Filename = MyName
            Selection.TypeText .Filename & vbTab & FileLen(MyName) _
               & vbTab & FileDateTime(MyName) & Chr$(13)
        Next i
        
        ' Type the total number of files found.
        Selection.TypeText Chr$(13)
        Selection.TypeText "Total files in folder = " & TotalFiles & _
           " files."
    
    End With
    
    Else
        MsgBox "The folder does not exist. Please try again."
        GoTo Folder
    End If
End If

PrintResponse = MsgBox("Do you want to print this folder list?", vbYesNo)
    If PrintResponse = "6" Then
        Application.ActiveDocument.PrintOut
    End If
    
AgainResponse = MsgBox("Do you want to list another folder?", vbYesNo)
    If AgainResponse = "6" Then
        GoTo Folder
    Else
        End
    End If

End:
End Sub



