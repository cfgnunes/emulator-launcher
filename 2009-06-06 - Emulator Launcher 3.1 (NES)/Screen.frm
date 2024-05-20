VERSION 5.00
Begin VB.Form Autorun 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Nintendo Games"
   ClientHeight    =   4950
   ClientLeft      =   3120
   ClientTop       =   150
   ClientWidth     =   8400
   ForeColor       =   &H00000000&
   Icon            =   "Screen.frx":0000
   LinkTopic       =   "Screen"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Screen.frx":000C
   ScaleHeight     =   330
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   560
   StartUpPosition =   2  'CenterScreen
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
   Begin VB.Image GameScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   3345
      Left            =   4290
      Stretch         =   -1  'True
      Top             =   795
      Width           =   3870
   End
   Begin VB.Label BtnRun 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   660
      Left            =   5145
      MouseIcon       =   "Screen.frx":875EE
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   4215
      Width           =   615
   End
   Begin VB.Label BtnExit 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   660
      Left            =   6690
      MouseIcon       =   "Screen.frx":87740
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   4215
      Width           =   615
   End
End
Attribute VB_Name = "Autorun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ******************************************
' Projeto: Emulator Launcher (for NES)
' Versão: 3.1
' Autor: Cristiano Fraga G. Nunes
' Data: 06/06/2009
' ******************************************

Option Explicit
Dim fso As New FileSystemObject
Dim ProgramPath As String
Dim ScreenFile As String
Dim x As Integer
Const ImagesDir = "Images"
Const ImagesExt = ".gif"
Const EmulatorDir = "Emulator"
Const EmulatorExe = "Nestopia.exe"
Const RomsDir = "Roms"
Const ConsoleNameDir = "nes"
Const DriveToInstall = "C:\"

Private Sub BtnExit_Click()
    ExitAutorun
End Sub

Private Sub BtnRun_Click()
    RunGame
End Sub

Private Sub Form_Load()
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Right$(App.Path, 1) = "\" Then
        ProgramPath = App.Path
    Else
        ProgramPath = App.Path + "\"
    End If
    CopyEmulator
    If fso.FolderExists(ProgramPath + RomsDir) = True Then RomList.Path = ProgramPath + RomsDir
    For x = 0 To RomList.ListCount - 1
        RomList.ListIndex = x
        GameList.AddItem Left$(RomList.FileName, Len(RomList.FileName) - 4)
    Next x
End Sub

Private Sub GameList_Click()
    DoEvents
    If fso.FileExists(ProgramPath + ImagesDir + "\" + GameList.Text + ImagesExt) Then
        GameScreen.Picture = LoadPicture(ProgramPath + ImagesDir + "\" + GameList.Text + ImagesExt)
    Else
        GameScreen.Picture = LoadPicture()
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
    If GameList.ListIndex >= 0 Then Shell (Chr$(34) + DriveToInstall + EmulatorDir + "\" + ConsoleNameDir + "\" + EmulatorExe + Chr$(34) + " " + Chr$(34) + ProgramPath + RomsDir + "\" + RomList.FileName + Chr$(34)), vbNormalFocus
End Sub

Sub CopyEmulator()
    On Error Resume Next
    If fso.FileExists(DriveToInstall + EmulatorDir + "\" + ConsoleNameDir + "\" + EmulatorExe) = False Then
        If fso.FolderExists(DriveToInstall + EmulatorDir) = False Then fso.CreateFolder (DriveToInstall + EmulatorDir)
        fso.CopyFolder ProgramPath + EmulatorDir, DriveToInstall + EmulatorDir + "\" + ConsoleNameDir, True
    End If
    On Error GoTo 0
End Sub
