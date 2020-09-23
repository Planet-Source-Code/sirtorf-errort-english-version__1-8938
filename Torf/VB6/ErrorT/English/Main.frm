VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "ErrorT"
   ClientHeight    =   7785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9630
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   9630
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CheckBox chkLog 
      Caption         =   "Show Logfile"
      Height          =   195
      Left            =   4680
      TabIndex        =   30
      Top             =   6600
      Value           =   1  'Aktiviert
      Width           =   1935
   End
   Begin VB.Frame Frame3 
      Caption         =   "If there is already a ""On Error..."""
      Enabled         =   0   'False
      Height          =   975
      Left            =   4680
      TabIndex        =   27
      Top             =   4440
      Width           =   4815
      Begin VB.OptionButton optLeaveOut 
         Caption         =   "Don't check this sub"
         Enabled         =   0   'False
         Height          =   195
         Left            =   240
         TabIndex        =   29
         Top             =   360
         Value           =   -1  'True
         Width           =   3375
      End
      Begin VB.OptionButton optReplace 
         Caption         =   "Replace it"
         Enabled         =   0   'False
         Height          =   195
         Left            =   240
         TabIndex        =   28
         Top             =   600
         Width           =   2655
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "On Error ..."
      Height          =   2295
      Left            =   4680
      TabIndex        =   18
      Top             =   2040
      Width           =   4815
      Begin VB.OptionButton optResume 
         Caption         =   "... Resume Next"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   2535
      End
      Begin VB.OptionButton optGoto0 
         Caption         =   "... Goto 0"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   600
         Width           =   2415
      End
      Begin VB.OptionButton optGotoErrorT 
         Caption         =   "... Goto ErrorT"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   840
         Value           =   -1  'True
         Width           =   2655
      End
      Begin RichTextLib.RichTextBox txtError 
         Height          =   855
         Left            =   480
         TabIndex        =   22
         Top             =   1320
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   1508
         _Version        =   393217
         ScrollBars      =   3
         TextRTF         =   $"Main.frx":030A
      End
      Begin VB.Label Label1 
         Caption         =   "ErrorT:"
         Height          =   255
         Left            =   480
         TabIndex        =   23
         Top             =   1080
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select which parts should be checked"
      Height          =   4815
      Left            =   120
      TabIndex        =   15
      Top             =   2040
      Width           =   4455
      Begin VB.CommandButton cmdSelectAll 
         Caption         =   "All"
         Height          =   375
         Left            =   1680
         TabIndex        =   25
         Top             =   4320
         Width           =   1215
      End
      Begin VB.CommandButton cmdSelectNone 
         Caption         =   "None"
         Height          =   375
         Left            =   3000
         TabIndex        =   24
         Top             =   4320
         Width           =   1215
      End
      Begin VB.OptionButton optAll 
         Caption         =   "Check all"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optSelected 
         Caption         =   "Check the parts which were selected"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   3255
      End
      Begin MSComctlLib.TreeView Tree 
         Height          =   3375
         Left            =   600
         TabIndex        =   26
         Top             =   840
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   5953
         _Version        =   393217
         Indentation     =   529
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         ImageList       =   "ImageList"
         Appearance      =   1
      End
   End
   Begin VB.CheckBox chkLineNumbers 
      Caption         =   "Write Linenumbers"
      Height          =   195
      Left            =   4680
      TabIndex        =   14
      Top             =   5520
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   6960
      TabIndex        =   11
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go !"
      Height          =   495
      Left            =   8280
      TabIndex        =   10
      Top             =   7200
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   5160
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Load project"
      Filter          =   "Visual Basic projects|*.vbp"
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   255
      Left            =   9000
      TabIndex        =   1
      Top             =   120
      Width           =   495
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   5160
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":03D3
            Key             =   "Projekt"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0827
            Key             =   "Modul"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0DBB
            Key             =   "Form"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":134F
            Key             =   "Sub"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":14B3
            Key             =   "Function"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label9 
      Caption         =   "SirTorf@web.de"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   8280
      TabIndex        =   33
      Top             =   6480
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "© 2000 by SirTorf"
      Height          =   255
      Left            =   8160
      TabIndex        =   32
      Top             =   6240
      Width           =   1335
   End
   Begin VB.Label Label7 
      Caption         =   "ErrorT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8280
      TabIndex        =   31
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   7680
      Top             =   5760
      Width           =   480
   End
   Begin VB.Label lblStatus 
      Height          =   255
      Left            =   600
      TabIndex        =   13
      Top             =   7320
      Width           =   5415
   End
   Begin VB.Label Label2 
      Caption         =   "Status:"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   7320
      Width           =   615
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   9480
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line2 
      X1              =   9480
      X2              =   120
      Y1              =   7080
      Y2              =   7080
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   9480
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label lblProjectFirma 
      Height          =   255
      Left            =   1440
      TabIndex        =   9
      Top             =   1200
      Width           =   3615
   End
   Begin VB.Label lblProjectCopyRight 
      Height          =   255
      Left            =   1440
      TabIndex        =   8
      Top             =   1440
      Width           =   3615
   End
   Begin VB.Label lblProjectVersion 
      Height          =   255
      Left            =   1440
      TabIndex        =   7
      Top             =   960
      Width           =   3615
   End
   Begin VB.Label lblProjectName 
      Height          =   255
      Left            =   1440
      TabIndex        =   6
      Top             =   720
      Width           =   3615
   End
   Begin VB.Label Label6 
      Caption         =   "Copyright:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Company:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Version:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Name of project:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label labProject 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8775
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'•••••••••••••••••••••
'• © 2000 by SirTorf •
'•  SirTorf@web.de   •
'•••••••••••••••••••••


Option Explicit

Private Sub chkLineNumbers_Click()
    SetLineNumbers = chkLineNumbers.Value
End Sub

Private Sub chkLog_Click()
    ShowLog = chkLog.Value
End Sub

Private Sub cmdBrowse_Click()
    LoadFile
End Sub

Private Sub cmdGo_Click()
    ImpfProject
End Sub

Private Sub cmdSelectAll_Click()
Dim a As Integer
    For a = 1 To Tree.Nodes.Count
        Tree.Nodes(a).Checked = True
    Next a
End Sub



Private Sub cmdSelectNone_Click()
Dim a As Integer
    For a = 1 To Tree.Nodes.Count
        Tree.Nodes(a).Checked = False
    Next a
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub


Private Function AddModuleNode(ByVal Formindex As Integer) As Node
Dim Picture As String
Dim Key As String
Dim Text As String
    If GetModuleType(CurrentProject.Forms(Formindex).Filename) = 1 Then
        Picture = "Form"
    ElseIf GetModuleType(CurrentProject.Forms(Formindex).Filename) = 2 Then
        Picture = "Modul"
    Else
        Picture = ""
    End If
    Key = CurrentProject.Forms(Formindex).Filename
    Text = CurrentProject.Forms(Formindex).Name
    Set AddModuleNode = Tree.Nodes.Add("project", tvwChild, Key, Text, Picture, Picture)
End Function

Private Function AddSubNode(ByVal Formindex As Integer, ByVal Subindex As Integer) As Node
Dim Picture As String
Dim Key As String
Dim Text As String
    If GetSubType(CurrentProject.Forms(Formindex).Subs(Subindex)) = 1 Then
        Picture = "Sub"
    ElseIf GetSubType(CurrentProject.Forms(Formindex).Subs(Subindex)) = 2 Then
        Picture = "Function"
    Else
        Picture = ""
    End If
    Key = CurrentProject.Forms(Formindex).Filename & "|" & _
                CurrentProject.Forms(Formindex).Subs(Subindex)
    Text = GetSubName(CurrentProject.Forms(Formindex).Subs(Subindex))
    Set AddSubNode = Tree.Nodes.Add(CurrentProject.Forms(Formindex).Filename, tvwChild, Key, _
                Text, Picture, Picture)
End Function


Private Sub FillTree()
Dim a As Integer
Dim b As Integer
Dim hNode As Node
    Tree.Nodes.Clear
    Set hNode = Tree.Nodes.Add(, , "project", CurrentProject.Name, "Projekt", "Projekt")
    hNode.Bold = True
    If UBound(CurrentProject.Forms) = 0 Then Exit Sub
    For a = 1 To UBound(CurrentProject.Forms)
        Set hNode = AddModuleNode(a)
        hNode.Bold = True
        If Not UBound(CurrentProject.Forms(a).Subs) = 0 Then
            For b = 1 To UBound(CurrentProject.Forms(a).Subs)
                AddSubNode a, b
            Next b
        Else
            Set hNode = GetNodeByKey(CurrentProject.Forms(a).Filename)
            hNode.Bold = False
        End If
    Next a
End Sub




Public Function GetNodePosition(ByVal sNode As Node) As Integer
Dim PNode As Node
Dim a As Integer
    On Error GoTo PosError
    If sNode.Index = 1 Then
        GetNodePosition = -1
        Exit Function
    End If
    If sNode.FirstSibling = sNode Then
        GetNodePosition = 1
        Exit Function
    End If
    Set PNode = sNode.FirstSibling
    a = 1
    Do Until PNode.Text = sNode.Text Or (PNode Is Nothing)
        a = a + 1
        Set PNode = PNode.Next
    Loop
    If PNode Is Nothing Then
        GetNodePosition = 0
    Else
        GetNodePosition = a
    End If
    Exit Function
PosError:
    MsgBox Str(Err.Number) & " " & Err.Description
End Function



Private Sub Form_Load()
    IsProjectLoaded = False
    LeaveOut = True
    ShowLog = True
    Dialog.Flags = cmdlFlagsOpen
    EnableAll False
    SetLineNumbers = False
    CheckAll = True
    StringAtStart = "On Error Goto ErrorT"
    StringAtEnd = "If Err.Number <> 0 Then" & vbCrLf & "    MsgBox ""Error!""" & vbCrLf & _
                "End If"
    txtError.Text = StringAtEnd
    Me.Show
    DoEvents
    LoadFile
Cancel:
    Status "Ready"
End Sub

Public Sub EnableAll(ByVal Yes As Boolean)
    Frame1.Enabled = Yes
    Frame2.Enabled = Yes
    chkLineNumbers.Enabled = Yes
    optResume.Enabled = Yes
    optGoto0.Enabled = Yes
    optGotoErrorT.Enabled = Yes
    txtError.Enabled = optGotoErrorT.Value And optGotoErrorT.Enabled
    optAll.Enabled = Yes
    optSelected.Enabled = Yes
    cmdSelectAll.Enabled = optSelected.Value
    cmdSelectNone.Enabled = optSelected.Value
    Tree.Enabled = optSelected.Value
    Label1.Enabled = Yes
    Label3.Enabled = Yes
    Label4.Enabled = Yes
    Label5.Enabled = Yes
    Label6.Enabled = Yes
    cmdGo.Enabled = Yes
End Sub

Private Sub LoadFile()
Dim a As Integer
Dim Filename As String
On Error GoTo OpenError
    If Not IsProjectLoaded Then EnableAll False
    Status "Waiting for your input..."
    Dialog.ShowOpen
    Filename = Dialog.Filename
    labProject.Caption = Dialog.Filename
    Screen.MousePointer = 11
    Status "Loading project..."
    LoadProject Filename
    FillTree
    lblProjectName.Caption = CurrentProject.Name
    lblProjectFirma.Caption = CurrentProject.Company
    lblProjectCopyRight.Caption = CurrentProject.CopyRight
    lblProjectVersion.Caption = Trim(Str(CurrentProject.MajorVersion)) & "." & _
                Trim(Str(CurrentProject.MinorVersion))
    Screen.MousePointer = 0
    Status "Bereit..."
    EnableAll True
OpenError:
    If Err.Number = 0 Then Exit Sub
    Status "Ready"
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload frmLog
End Sub

Private Sub Label9_Click()
    ShellExecute Me.hwnd, vbNullString, "mailto: SirTorf@web.de", "vbnullstring", "C:\", _
                SW_SHOWNORMAL
End Sub

Private Sub optAll_Click()
    CheckAll = optAll.Value
    Tree.Enabled = Not CheckAll
    cmdSelectAll.Enabled = optSelected.Value
    cmdSelectNone.Enabled = optSelected.Value
End Sub

Private Sub optGoto0_Click()
    StringAtStart = "On Error Goto 0"
    txtError.Enabled = False
End Sub

Private Sub optGotoErrorT_Click()
    StringAtStart = "On Error Goto ErrorT"
    txtError.Enabled = True
End Sub

Private Sub optLeaveOut_Click()
    LeaveOut = True
End Sub

Private Sub optReplace_Click()
    LeaveOut = False
End Sub

Private Sub optResume_Click()
    StringAtStart = "On Error Resume Next"
    txtError.Enabled = False
End Sub

Private Sub optSelected_Click()
    CheckAll = Not optSelected.Value
    Tree.Enabled = Not CheckAll
    cmdSelectAll.Enabled = optSelected.Value
    cmdSelectNone.Enabled = optSelected.Value
End Sub


Private Sub txtError_Change()
    StringAtEnd = txtError.Text
End Sub
