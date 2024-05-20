VERSION 5.00
Begin VB.Form Autorun 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   5550
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8400
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   Icon            =   "Screen.frx":0000
   LinkTopic       =   "Screen"
   LockControls    =   -1  'True
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   370
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   560
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox GameScreen 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   3345
      Left            =   4290
      ScaleHeight     =   221
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   795
      Width           =   3870
      Begin VB.Label lblScreen 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   7
         Top             =   1560
         Width           =   3855
      End
   End
   Begin VB.ListBox GameList 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   3345
      Left            =   240
      TabIndex        =   1
      Top             =   795
      Width           =   3870
   End
   Begin VB.FileListBox RomList 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   420
      Left            =   240
      Pattern         =   "*.zip"
      TabIndex        =   0
      Top             =   795
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label btnExit 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4305
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Label btnRun 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2760
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   5190
      Width           =   7815
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   7815
   End
End
Attribute VB_Name = "Autorun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ******************************************
' Projeto: Emulator Launcher
' Versão: 4.0
' Autor: Cristiano Fraga G. Nunes
' Data: 23/06/2005
' ******************************************

Option Explicit
Dim ProgramPath As String
Dim ScreenFile As String
Dim xInt As Integer
Dim EmulatorDir As String
Dim ImagesDir As String
Dim RomsDir As String
Dim ImagesExt As String
Dim EmulatorExe As String

Private Sub BtnExit_Click()
    ExitAutorun
End Sub

Private Sub BtnRun_Click()
    RunGame
End Sub

Private Sub Form_Load()
    DrawBackground
    LoadResource
    If Right$(App.Path, 1) = Chr$(92) Then
        ProgramPath = App.Path
    Else
        ProgramPath = App.Path + Chr$(92)
    End If
    If Dir(ProgramPath + RomsDir, vbDirectory) <> "" Then
        RomList.Path = ProgramPath + RomsDir
        For xInt = 0 To RomList.ListCount - 1
            RomList.ListIndex = xInt
            GameList.AddItem Left$(RomList.FileName, Len(RomList.FileName) - (Len(RomList.Pattern) - 1))
        Next xInt
'        lblStatus.Caption = Trim$(Str$(GameList.ListCount)) + " " + LoadResString(16)
    End If
    If GameList.ListCount > 0 Then GameList.ListIndex = 0
    lblStatus.Caption = "Autor: Cristiano Fraga G. Nunes"
    Autorun.Visible = True
End Sub

Private Sub GameList_Click()
    DoEvents
    If Dir(ProgramPath + ImagesDir + Chr$(92) + GameList.Text + ImagesExt) <> "" Then
        lblScreen.Visible = False
        GameScreen.Picture = LoadPicture(ProgramPath + ImagesDir + Chr$(92) + GameList.Text + ImagesExt)
    Else
        lblScreen.Caption = LoadResString(22)
        GameScreen.Picture = LoadPicture()
        lblScreen.Visible = True
    End If
End Sub

Private Sub GameList_KeyPress(KeyAscii As Integer)
    If KeyAscii = Val(vbKeyEscape) Then ExitAutorun
    If KeyAscii = Val(vbKeyReturn) Then RunGame
End Sub

Sub ExitAutorun()
    Unload Me
    End
End Sub

Sub RunGame()
    RomList.ListIndex = GameList.ListIndex
    If GameList.ListIndex >= 0 Then
        If Dir(ProgramPath + EmulatorDir + Chr$(92) + EmulatorExe) <> "" Then
            ChDir ProgramPath + EmulatorDir
            Shell (Chr$(34) + ProgramPath + EmulatorDir + Chr$(92) + EmulatorExe + Chr$(34) + Chr$(32) + Chr$(34) + ProgramPath + RomsDir + Chr$(92) + RomList.FileName + Chr$(34)), vbNormalFocus
        Else
            MsgBox LoadResString(21), vbOKOnly + vbCritical, LoadResString(6)
        End If
    End If
End Sub

Sub LoadResource()
    EmulatorDir = LoadResString(0)
    ImagesDir = LoadResString(1)
    RomsDir = LoadResString(2)
    ImagesExt = LoadResString(3)
    EmulatorExe = LoadResString(5)
    RomList.Pattern = "*" + LoadResString(4)
    Autorun.Caption = LoadResString(6)
    lblTitle.Caption = LoadResString(6)
    btnRun.Caption = LoadResString(17)
    btnExit.Caption = LoadResString(18)
    btnRun.ToolTipText = LoadResString(19)
    btnExit.ToolTipText = LoadResString(20)
    btnRun.MouseIcon = LoadResPicture(1, 2)
    btnExit.MouseIcon = LoadResPicture(1, 2)
End Sub

Sub DrawBackground()
    Dim rBk, bBk, gBk, brBk, pBk As Integer
    
    GoSub ResetaCores
    pBk = 0
    For xInt = 100 To 330
        Circle (Autorun.ScaleWidth / 2 + 1, Autorun.ScaleHeight / 2), xInt, RGB(rBk, gBk, bBk)
        Circle (Autorun.ScaleWidth / 2 - 1, Autorun.ScaleHeight / 2), xInt, RGB(rBk, gBk, bBk)
        Circle (Autorun.ScaleWidth / 2, Autorun.ScaleHeight / 2 + 1), xInt, RGB(rBk, gBk, bBk)
        Circle (Autorun.ScaleWidth / 2, Autorun.ScaleHeight / 2 - 1), xInt, RGB(rBk, gBk, bBk)
        Circle (Autorun.ScaleWidth / 2, Autorun.ScaleHeight / 2), xInt, RGB(rBk, gBk, bBk)
        pBk = pBk + 1
        If pBk = 8 Then
            GoSub DegradeCores
            pBk = 0
        End If
    Next xInt
    
    GoSub ResetaCores
    For xInt = 30 To 0 Step -1
        Line (0, xInt)-(Autorun.ScaleWidth, xInt), RGB(rBk, gBk, bBk)
        GoSub DegradeCores
    Next xInt
    
    GoSub ResetaCores
    For xInt = 0 To 30
        Line (0, xInt + (Autorun.ScaleHeight - 30 - 1))-(Autorun.ScaleWidth, xInt + (Autorun.ScaleHeight - 30 - 1)), RGB(rBk, gBk, bBk)
        GoSub DegradeCores
    Next xInt
    
    GoSub ResetaCores
    For xInt = 0 To btnExit.Height - 1
        Line (btnExit.Left, xInt + btnExit.Top)-(btnExit.Left + btnExit.Width, xInt + btnExit.Top), RGB(rBk, gBk, bBk)
        GoSub DegradeCores
    Next xInt
    
    GoSub ResetaCores
    For xInt = 0 To btnRun.Height - 1
        Line (btnRun.Left, xInt + btnRun.Top)-(btnRun.Left + btnRun.Width, xInt + btnRun.Top), RGB(rBk, gBk, bBk)
        GoSub DegradeCores
    Next xInt
    
    Exit Sub
    
ResetaCores:
    rBk = 255
    gBk = 255
    bBk = 255
    Return
    
DegradeCores:
    rBk = rBk - 3
    gBk = gBk - 2
    bBk = bBk - 1
    If rBk <= 0 Then rBk = 0
    If gBk <= 0 Then gBk = 0
    If bBk <= 0 Then bBk = 0
    Return
End Sub
