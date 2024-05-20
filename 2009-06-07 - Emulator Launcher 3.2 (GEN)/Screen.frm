VERSION 5.00
Begin VB.Form Autorun 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Sega Genesis Games"
   ClientHeight    =   4950
   ClientLeft      =   3120
   ClientTop       =   150
   ClientWidth     =   8400
   ForeColor       =   &H00000000&
   Icon            =   "screen.frx":0000
   LinkTopic       =   "Screen"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "screen.frx":000C
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
      MouseIcon       =   "screen.frx":875EE
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
      MouseIcon       =   "screen.frx":87740
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
' Projeto: Emulator Launcher (for Genesis)
' Versão: 3.2
' Autor: Cristiano Fraga G. Nunes
' Data: 07/06/2009
' ******************************************

Option Explicit
Dim ObjFso As New FileSystemObject
Dim ObjShell As Object

Dim AppData As String
Dim ProgramPath As String
Dim ScreenFile As String
Dim x As Integer

Const ImagesDir = "images"
Const ImagesExt = ".gif"
Const EmulatorDir = "emulator"
Const EmulatorExe = "fusion.exe"
Const RomsDir = "roms"
Const ConsoleNameDir = "gen"

Private Sub BtnExit_Click()
    ExitAutorun
End Sub

Private Sub BtnRun_Click()
    RunGame
End Sub

Private Sub Form_Load()
    Set ObjFso = CreateObject("Scripting.FileSystemObject")
    Set ObjShell = CreateObject("WScript.Shell")
    AppData = ObjShell.ExpandEnvironmentStrings("%APPDATA%") & "\"
    If Right$(App.Path, 1) = "\" Then
        ProgramPath = App.Path
    Else
        ProgramPath = App.Path & "\"
    End If
    CopyEmulator
    If ObjFso.FolderExists(ProgramPath & RomsDir) = True Then RomList.Path = ProgramPath & RomsDir
    For x = 0 To RomList.ListCount - 1
        RomList.ListIndex = x
        GameList.AddItem Left$(RomList.FileName, Len(RomList.FileName) - 4)
    Next x
End Sub

Private Sub GameList_Click()
    DoEvents
    If ObjFso.FileExists(ProgramPath & ImagesDir & "\" & GameList.Text & ImagesExt) Then
        GameScreen.Picture = LoadPicture(ProgramPath & ImagesDir & "\" & GameList.Text & ImagesExt)
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
    ChangeDir AppData & EmulatorDir & "\" & ConsoleNameDir & "\"
    If GameList.ListIndex >= 0 Then Shell (Chr$(34) & AppData & EmulatorDir & "\" & ConsoleNameDir & "\" & EmulatorExe & Chr$(34) & " " & Chr$(34) & ProgramPath & RomsDir & "\" & RomList.FileName & Chr$(34)), vbNormalFocus
End Sub

Sub CopyEmulator()
    On Error Resume Next
    If ObjFso.FileExists(AppData & EmulatorDir & "\" & ConsoleNameDir & "\" & EmulatorExe) = False Then
        If ObjFso.FolderExists(AppData & EmulatorDir) = False Then ObjFso.CreateFolder (AppData & EmulatorDir)
        ObjFso.CopyFolder ProgramPath & EmulatorDir, AppData & EmulatorDir & "\" & ConsoleNameDir, True
    End If
    On Error GoTo 0
End Sub

Sub ChangeDir(Path As String)
    Dim TargetDrive As String
    If Mid(Path, 2, 2) = ":\" Then
        TargetDrive = Left(Path, 3)
        If TargetDrive <> Left(CurDir, 3) Then
            ChDrive TargetDrive
        End If
    End If
    VBA.ChDir Path
End Sub
