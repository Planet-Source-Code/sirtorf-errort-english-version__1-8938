Attribute VB_Name = "ErrModule"
'•••••••••••••••••••••
'• © 2000 by SirTorf •
'•  SirTorf@web.de   •
'•••••••••••••••••••••

Option Explicit

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
            ByVal hwnd As Long, _
            ByVal lpOperation As String, _
            ByVal lpFile As String, _
            ByVal lpParameters As String, _
            ByVal lpDirectory As String, _
            ByVal nShowCmd As Long) As Long
Public Const SW_SHOWNORMAL = 1

Public Type xForm
    Name As String
    Filename As String
    Subs() As String
End Type

Public Type Project
    Path As String
    Name As String
    Version As String
    MajorVersion As Integer
    MinorVersion As Integer
    RevisionVersion As Integer
    Company As String
    CopyRight As String
    Forms() As xForm
    Filename As String
End Type

Public Const cmdlFlagsOpen = cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNLongNames Or _
                             cdlOFNNoChangeDir
                             
Public Const cmdlFlagsSave = cdlOFNNoChangeDir Or cdlOFNHideReadOnly Or cdlOFNLongNames Or _
                             cdlOFNOverwritePrompt


Public Const cmdlFlagsColor = cdlCCFullOpen



Public CurrentProject As Project
Public StringAtStart As String
Public StringAtEnd As String
Public CheckAll As Boolean
Public SetLineNumbers As Boolean
Public LeaveOut As Boolean
Public IsProjectLoaded As Boolean
Public ShowLog As Boolean

Private SubsImpfed As Integer
Private FormsImpfed As Integer
Private ZeilenGesamt As Long
Private NewFolder As String


Public Function GetFormFile(ByVal Text As String) As String
Dim Pos As Integer
    Pos = InStr(Text, "=")
    If Pos = 0 Then
        GetFormFile = Trim(Text)
    Else
        GetFormFile = Trim(Right(Text, Len(Text) - Pos))
    End If
    GetFormFile = CheckPath(GetFormFile)
End Function

Public Function GetModuleFile(ByVal Text As String) As String
Dim Pos As Integer
Dim Help As String
    Pos = InStr(Text, ";")
    If Pos = 0 Then
        GetModuleFile = Trim(Text)
    Else
        GetModuleFile = Trim(Right(Text, Len(Text) - Pos))
    End If
    GetModuleFile = CheckPath(GetModuleFile)
End Function

Public Sub LoadProject(ByVal Filename As String)
Dim Help(1 To 1000) As String
Dim Filenum As Integer
Dim Zeile As String
Dim a As Integer
Dim b As Integer
    Filenum = FreeFile
    a = 0
    CurrentProject.Path = GetProjectPath(Filename)
    CurrentProject.Filename = Filename
    Open Filename For Input As #Filenum
        Do While Not EOF(Filenum)
            Line Input #Filenum, Zeile
            If UCase(Left(Zeile, 4)) = "FORM" Then
                a = a + 1
                Help(a) = GetFormFile(Zeile)
            ElseIf UCase(Left(Zeile, 6)) = "MODULE" Then
                a = a + 1
                Help(a) = GetModuleFile(Zeile)
            ElseIf UCase(Left(Zeile, 4)) = "NAME" Then
                CurrentProject.Name = GetValue(Zeile)
            ElseIf UCase(Left(Zeile, 8)) = "MAJORVER" Then
                CurrentProject.MajorVersion = Val(GetValue(Zeile))
            ElseIf UCase(Left(Zeile, 8)) = "MINORVER" Then
                CurrentProject.MinorVersion = Val(GetValue(Zeile))
            ElseIf UCase(Left(Zeile, 18)) = "VERSIONCOMPANYNAME" Then
                CurrentProject.Company = GetValue(Zeile)
            ElseIf UCase(Left(Zeile, 21)) = "VERSIONLEGALCOPYRIGHT" Then
                CurrentProject.CopyRight = GetValue(Zeile)
            End If
            DoEvents
        Loop
    Close #Filenum
    ReDim CurrentProject.Forms(1 To a)
    For b = 1 To a
        CurrentProject.Forms(b).Filename = Help(b)
    Next b
    For b = 1 To a
        LoadSubs b
    Next b
    IsProjectLoaded = True
End Sub

Public Function GetModuleType(ByVal Filename As String) As Byte
    If UCase(Right(Filename, 3)) = "FRM" Then
        GetModuleType = 1
    ElseIf UCase(Right(Filename, 3)) = "BAS" Then
        GetModuleType = 2
    Else
        GetModuleType = 0
    End If
End Function

Public Sub LoadSubs(ByVal Formindex As Integer)
Dim a As Integer
Dim b As Integer
Dim Help(1 To 1000) As String
Dim Filenum As Integer
Dim Zeile As String
    Filenum = FreeFile
    a = 0
    Open CurrentProject.Forms(Formindex).Filename For Input As #Filenum
        Do While Not EOF(Filenum)
            Line Input #Filenum, Zeile
            If IsStartOfSub(Zeile) Then
                a = a + 1
                Help(a) = Trim(Zeile)
            ElseIf Trim(Left(Zeile, 17)) = "Attribute VB_Name" Then
                CurrentProject.Forms(Formindex).Name = GetValue(Zeile)
            End If
        Loop
    Close #Filenum
    If a = 0 Then
        ReDim CurrentProject.Forms(Formindex).Subs(0)
        Exit Sub
    End If
    ReDim CurrentProject.Forms(Formindex).Subs(1 To a)
    For b = 1 To a
        CurrentProject.Forms(Formindex).Subs(b) = Help(b)
    Next b
    Status "Ready"
End Sub

Public Function GetProjectPath(ByVal Fullname As String) As String
Dim Pos As Integer
    Pos = InStrRev(Fullname, "\")
    If Pos = 0 Then
        GetProjectPath = EnsureBackslash(Fullname)
    Else
        GetProjectPath = EnsureBackslash(Trim(Left(Fullname, Pos)))
    End If
End Function

Private Function GetValue(ByVal Text As String) As String
Dim Pos As Integer
    Pos = InStr(Text, "=")
    If Pos = 0 Then
        GetValue = Text
    Else
        GetValue = Right(Text, Len(Text) - Pos)
    End If
End Function

Private Function ReadAndWriteHeader(ByVal FilenumR As Integer, ByVal FilenumW As Integer) As String
Dim Zeile As String
    Status "Reading and writing header..."
    Do While UCase((Left(Zeile, 12))) <> "ATTRIBUTE VB"
        Line Input #FilenumR, Zeile
        Print #FilenumW, Zeile
    Loop
    Do While UCase((Left(Zeile, 12))) = "ATTRIBUTE VB"
        Print #FilenumW, Zeile
        Line Input #FilenumR, Zeile
    Loop
    ReadAndWriteHeader = Zeile
    Status "Bereit"
End Function

Public Function EnsureBackslash(ByVal Path As String) As String
    If Right(Path, 1) <> "\" Then Path = Path + "\"
    EnsureBackslash = Path
End Function

Public Sub ImpfProject()
Dim a As Integer
Dim b As Integer
Dim hNode As Node
Dim Fso As New FileSystemObject
Dim Zeit As Date
    Zeit = Now
    If ShowLog Then frmLog.Show
    DoEvents
    Log "---Checking: " & Trim(CurrentProject.Filename)

    ZeilenGesamt = 0
    SubsImpfed = 0
    FormsImpfed = 0
    
    If frmMain.Tree.Nodes(1).Checked = False And Not CheckAll Then
        Log "Project wasn't checked"
        Log "---Finished checking: " & Trim(CurrentProject.Filename)
        Exit Sub
    End If
    
    If UBound(CurrentProject.Forms) = 0 Then
        Log "Keine Forms vorhanden, Projekt wird nicht geimpft"
        Log "---Finished checking: " & Trim(CurrentProject.Filename)
        Exit Sub
    End If
    
    NewFolder = EnsureBackslash(CurrentProject.Path & "Checked version of " & _
                Trim(GetFilename(CurrentProject.Filename)))
    If Fso.FolderExists(Left(NewFolder, Len(NewFolder) - 1)) Then
        Log "Folder exists already - renaming"

        Name Left(NewFolder, Len(NewFolder) - 1) As Left(NewFolder, Len(NewFolder) - 1) & _
            " " & CStr(Format$(Now, "DD.MM.YY HH.MM.SS"))
        
    End If
    
    MkDir Left(NewFolder, Len(NewFolder) - 1)
    
    For a = 1 To UBound(CurrentProject.Forms)
        Set hNode = GetNodeByKey(CurrentProject.Forms(a).Filename)
        If (hNode.Checked Or CheckAll) Then
            SetErrorToSubs (a)
            FormsImpfed = FormsImpfed + 1
            Log "Form checked: " & Trim(CurrentProject.Forms(a).Filename)
        Else
            Log "Form wasn't checked: " & Trim(CurrentProject.Forms(a).Filename)
        End If
    Next a
    Log "---Finished checking: " & Trim(CurrentProject.Filename)
    ShowResult DateDiff("S", Zeit, Now)
End Sub



Public Sub SetErrorToSubs(ByVal Formindex As Integer)
Dim FilenumR As Integer
Dim FilenumW As Integer
Dim Help As String
Dim DontSet As Boolean
Dim Dummy As String
Dim LineCount As Long
Dim hNode As Node
Dim Subname As String
Dim NextLine As Boolean
    LineCount = 0
    FilenumR = FreeFile
    DontSet = False
    Help = CurrentProject.Forms(Formindex).Filename
    Open CurrentProject.Forms(Formindex).Filename For Input As #FilenumR
    FilenumW = FreeFile
    Open NewFolder & GetFilename(CurrentProject.Forms(Formindex).Filename) For Output As FilenumW
    Help = ReadAndWriteHeader(FilenumR, FilenumW)
    Status "Reading and writing subs..."
    Do
        ZeilenGesamt = ZeilenGesamt + 1
        If IsStartOfSub(Help) Then
            Print #FilenumW, Help
            Subname = Help
            If Right(Trim(Help), 1) = "_" Then
                    Line Input #FilenumR, Dummy
                    Print #FilenumW, Dummy
                    NextLine = True
            End If
            Line Input #FilenumR, Dummy
            Do While IsDim(Dummy) Or IsError(Dummy) Or Right(Trim(Dummy), 1) = "_"
                If Right(Trim(Dummy), 1) = "_" Then
                    Line Input #FilenumR, Dummy
                    Print #FilenumW, Dummy
                    NextLine = True
                End If
                If IsError(Dummy) Then
                    DontSet = True
                    Log """On Error """ & " already placed: " & _
                                Trim(CurrentProject.Forms(Formindex).Filename) & _
                                " / " & Subname
                End If
                Print #FilenumW, Dummy
                
                Line Input #FilenumR, Dummy
            Loop
            Set hNode = GetNodeByKey(CurrentProject.Forms(Formindex).Filename & "|" & Subname)
            If hNode Is Nothing Then
                DontSet = True
                Log "Error: Couldn't find entry - " & Subname
            Else
                If (hNode.Checked Or CheckAll) Then
                    Print #FilenumW, StringAtStart
                    Log "Sub checked: " & Trim(CurrentProject.Forms(Formindex).Filename) & _
                                " / " & Subname
                Else
                    DontSet = True
                    Log "Sub wasn't checked: " & Trim(CurrentProject.Forms(Formindex).Filename) & _
                                " / " & Subname
                End If
            End If
            Help = Dummy
            GoTo PrintElse
        ElseIf IsEndOfSub(Help) Then
            If DontSet Then
                DontSet = False
                Print #FilenumW, Help
                GoTo NextLoop
            ElseIf UCase(StringAtStart) = "ON ERROR GOTO 0" Then
                Print #FilenumW, Help
                GoTo NextLoop
            Else
                Print #FilenumW, "ErrorT:"
                Print #FilenumW, StringAtEnd
                Print #FilenumW, Help
                SubsImpfed = SubsImpfed + 1
            End If
            
        Else
PrintElse:
            If SetLineNumbers And Not IsReserved(Help) And Not Trim(Help) = "" Then
                LineCount = LineCount + 10
                Print #FilenumW, GetLineNumber(LineCount) & Help
            Else
                Print #FilenumW, Help
            End If
        End If
NextLoop:
    If Right(Trim(Help), 1) = "_" And Not NextLine Then
        Line Input #FilenumR, Help
        Do While Right(Trim(Help), 1) = "_"
            Print #FilenumW, Help
            Line Input #FilenumR, Help
        Loop
        Print #FilenumW, Help
    End If
    NextLine = False
    If EOF(FilenumR) Then Exit Do
    Line Input #FilenumR, Help
    Loop
    Close #FilenumR
    Close #FilenumW
    Status "Ready"
End Sub

Private Sub ShowResult(ByVal Time As Long)
Dim Msg As String
    Msg = "Result:" & vbCrLf & _
          vbCrLf & _
           Trim(Str(FormsImpfed)) & " File(s) and" & vbCrLf & _
          Trim(Str(SubsImpfed)) & " sub(s)." & vbCrLf & _
          "     were/was checked at all" & vbCrLf & vbCrLf & _
          "This part of the project consists of " & Trim(Str(ZeilenGesamt)) & _
                    " lines" & vbCrLf & vbCrLf & _
          "Duration: " & Trim(Str(Time)) & "s"
          
    MsgBox Msg, vbOKOnly + vbInformation, "Result"
End Sub

Private Function GetLineNumber(ByVal Number As Integer) As String
    GetLineNumber = Trim(Str(Number)) & Space(6 - Len(Trim(Str(Number))))
End Function

Private Function IsStartOfSub(ByVal Text As String) As Boolean
    IsStartOfSub = False
    If (UCase(Left(Trim(Text), 6)) = "PUBLIC") Or _
               (UCase(Left(Trim(Text), 7)) = "PRIVATE") Then
                    If UCase(Mid(Trim(Text), 8, 4)) = "SUB " Or _
                       UCase(Mid(Trim(Text), 8, 9)) = "FUNCTION " Or _
                       UCase(Mid(Trim(Text), 9, 4)) = "SUB " Or _
                       UCase(Mid(Trim(Text), 9, 9)) = "FUNCTION " Then
                            IsStartOfSub = True
                    End If
    ElseIf (UCase(Left(Trim(Text), 4))) = "SUB " Or _
           (UCase(Left(Trim(Text), 9))) = "FUNCTION " Then
            IsStartOfSub = True
    End If
End Function

Private Function IsEndOfSub(ByVal Text As String) As Boolean
    Text = Trim(Text)
    IsEndOfSub = False
    If UCase(Text) = "END SUB" Or UCase(Text) = "END FUNCTION" Then IsEndOfSub = True
End Function

Private Function IsDim(ByVal Text As String) As Boolean
    Text = Trim(Text)
    IsDim = False
    IsDim = UCase(Left(Text, 4)) = "DIM " Or UCase(Left(Text, 7)) = "STATIC " Or _
       (UCase(Left(Text, 6)) = "PUBLIC" And Not IsStartOfSub(Text)) Or _
       (UCase(Left(Text, 7)) = "PRIVATE" And Not IsStartOfSub(Text)) Or _
       UCase(Left(Text, 5)) = "CONST"
End Function

Private Function IsDefinition(ByVal Text As String) As Boolean
    Text = Trim(Text)
    IsDefinition = False
    If UCase(Left(Text, 3)) <> "DEF" Then Exit Function
    Text = Right(Text, Len(Text) - 3)
    IsDefinition = UCase(Left(Text, 4)) = "BOOL" Or UCase(Left(Text, 4)) = "DATE" Or _
                 UCase(Left(Text, 4)) = "BYTE" Or UCase(Left(Text, 3)) = "INT" Or _
                 UCase(Left(Text, 3)) = "LNG" Or UCase(Left(Text, 3)) = "CUR" Or _
                 UCase(Left(Text, 3)) = "SNG" Or UCase(Left(Text, 3)) = "DBL" Or _
                 UCase(Left(Text, 3)) = "DEC" Or UCase(Left(Text, 3)) = "STR" Or _
                 UCase(Left(Text, 3)) = "OBJ" Or UCase(Left(Text, 3)) = "VAR"
End Function

Private Function IsReserved(ByVal Text As String) As Boolean
    Text = Trim(Text)
    IsReserved = IsDim(Text) Or IsStartOfSub(Text) Or IsEndOfSub(Text) Or IsError(Text) Or _
            IsDefinition(Text) Or UCase(Left(Text, 6)) = "OPTION" Or _
            UCase(Left(Text, 7)) = "DECLARE" Or Left(Text, 1) = Chr(39)
    
End Function

Private Function IsError(ByVal Text As String) As Boolean
    Text = Trim(Text)
    IsError = UCase(Left(Text, 9)) = "ON ERROR "
End Function

Private Function CheckPath(ByVal Path As String) As String
Dim Pos As Integer
Dim HP As String
    HP = CurrentProject.Path
    If Right(HP, 1) = "\" Then HP = Left(HP, Len(HP) - 1)
    Path = Trim(Path)
    If Left(Path, 2) = ".." Then
        Pos = InStrRev(HP, "\")
        If Pos = 0 Then
            CheckPath = CurrentProject.Path + Right(Path, Len(Path) - 3)
        Else
            HP = Left(HP, Pos)
            CheckPath = HP + Right(Path, Len(Path) - 3)
        End If
    Else
        CheckPath = CurrentProject.Path + Path
    End If
End Function

Public Function GetSubNamePro(ByVal Text As String) As String
Dim Pos As Integer
    Pos = InStrRev(Text, "(")
    If Pos = 0 Then
        GetSubNamePro = Text
    Else
        GetSubNamePro = Left(Text, Pos - 1)
    End If
End Function

Public Function GetSubType(ByVal Text As String) As Byte
    Text = Trim(Text)
    If UCase(Left(Text, 7)) = "PUBLIC " Then
        Text = Right(Text, Len(Text) - 7)
    ElseIf UCase(Left(Text, 8)) = "PRIVATE " Then
        Text = Right(Text, Len(Text) - 8)
    End If
    If UCase(Left(Text, 4)) = "SUB " Then
        GetSubType = 1
    ElseIf UCase(Left(Text, 9)) = "FUNCTION " Then
        GetSubType = 2
    Else
        GetSubType = 0
    End If
End Function

Public Function GetSubName(ByVal Text As String) As String
    Text = Trim(Text)
    If UCase(Left(Text, 7)) = "PUBLIC " Then
        Text = Right(Text, Len(Text) - 7)
    ElseIf UCase(Left(Text, 8)) = "PRIVATE " Then
        Text = Right(Text, Len(Text) - 8)
    End If
    If UCase(Left(Text, 4)) = "SUB " Then
        Text = Right(Text, Len(Text) - 4)
        Text = Left(Text, InStr(Text, "(") - 1)
    ElseIf UCase(Left(Text, 9)) = "FUNCTION " Then
        Text = Right(Text, Len(Text) - 9)
        Text = Left(Text, InStr(Text, "(") - 1)
    End If
    GetSubName = Trim(Text)
End Function

Public Sub Status(ByVal Text As String)
    frmMain.lblStatus = " " & Text
    frmMain.lblStatus.Refresh
End Sub

Public Function GetNodeByKey(ByVal Key As String) As Node
Dim hNode As Node
Dim a As Integer
    For a = 1 To frmMain.Tree.Nodes.Count
        Set hNode = frmMain.Tree.Nodes(a)
        If hNode.Key = Key Then
            Set GetNodeByKey = hNode
            Exit Function
        End If
    Next a
End Function

Public Sub Log(ByVal Text As String)
    If Not ShowLog Then Exit Sub
    frmLog.txtLog.Text = frmLog.txtLog.Text + vbCrLf + Text
    frmLog.txtLog.Refresh
End Sub

Private Function GetFilename(ByVal Pfad As String) As String
Dim Pos As Integer
    If Right(Pfad, 1) = "\" Then Pfad = Left(Pfad, Len(Pfad) - 1)
    Pos = InStrRev(Pfad, "\")
    If Pos = 0 Then
        GetFilename = Pfad
    Else
        GetFilename = Right(Pfad, Len(Pfad) - Pos)
    End If
End Function
