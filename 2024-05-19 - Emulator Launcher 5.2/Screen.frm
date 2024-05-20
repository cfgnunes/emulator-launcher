VERSION 5.00
Begin VB.Form Launcher 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Emulator Launcher"
   ClientHeight    =   4950
   ClientLeft      =   3165
   ClientTop       =   480
   ClientWidth     =   8400
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   Icon            =   "Screen.frx":0000
   LinkTopic       =   "Screen"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   Picture         =   "Screen.frx":000C
   ScaleHeight     =   330
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   560
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox GameList 
      Appearance      =   0  'Flat
      BackColor       =   &H00F2F0EF&
      ForeColor       =   &H00565554&
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
   Begin VB.Label BtnExit 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00565554&
      Height          =   330
      Left            =   6390
      MouseIcon       =   "Screen.frx":875EE
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   4380
      Width           =   1305
   End
   Begin VB.Label BtnRun 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00565554&
      Height          =   330
      Left            =   4755
      MouseIcon       =   "Screen.frx":87740
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   4380
      Width           =   1305
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00565554&
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   8175
   End
   Begin VB.Label lblStatusList 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00006B3D&
      Height          =   210
      Left            =   1035
      TabIndex        =   2
      Top             =   4440
      Width           =   2325
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
End
Attribute VB_Name = "Launcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ******************************************
' Projeto: Emulator Launcher
' Versão: 5.2
' Autor: Cristiano Fraga G. Nunes
' Data: 19/05/2024
' ******************************************

Option Explicit
Dim ObjFso As New FileSystemObject
Dim ObjShell As Object

Dim ProgramPath As String
Dim ScreenFile As String
Dim x As Integer

Const ImagesDir = "Screenshots"
Const EmulatorDir = "Emulator"
Const RomsDir = "Roms"
Const ConfigFile = "Launcher.cfg"

Dim cfgWindowTitle As String
Dim cfgEmulatorExe As String
Dim cfgNewEmulatorDir As String
Dim cfgRomsExt As String
Dim cfgImagesExt As String

Private Sub BtnExit_Click()
    ExitLauncher
End Sub

Private Sub BtnRun_Click()
    RunGame
End Sub

Private Sub Form_Load()
    Set ObjFso = CreateObject("Scripting.FileSystemObject")
    Set ObjShell = CreateObject("WScript.Shell")
    If Right$(App.Path, 1) = "\" Then
        ProgramPath = App.Path
    Else
        ProgramPath = App.Path & "\"
    End If
    ReadConfig
    cfgNewEmulatorDir = ObjShell.ExpandEnvironmentStrings(cfgNewEmulatorDir)
    If Right$(cfgNewEmulatorDir, 1) <> "\" Then cfgNewEmulatorDir = cfgNewEmulatorDir & "\"
    CopyEmulator
    If ObjFso.FolderExists(ProgramPath & RomsDir) = True Then RomList.Path = ProgramPath & RomsDir
    RomList.Pattern = "*." & cfgRomsExt
    For x = 0 To RomList.ListCount - 1
        RomList.ListIndex = x
        GameList.AddItem Left$(RomList.FileName, Len(RomList.FileName) - 4)
    Next x
    Launcher.Caption = cfgWindowTitle
    lblTitle.Caption = cfgWindowTitle
    lblStatusList.Caption = Trim(Str(GameList.ListCount)) & " games found!"
End Sub

Private Sub GameList_Click()
    DoEvents
    If ObjFso.FileExists(ProgramPath & ImagesDir & "\" & GameList.Text & "." & cfgImagesExt) Then
        GameScreen.Picture = LoadPicture(ProgramPath & ImagesDir & "\" & GameList.Text & "." & cfgImagesExt)
    Else
        GameScreen.Picture = LoadPicture()
    End If
    lblStatusList.Caption = "Game " & Trim(Str(GameList.ListIndex + 1)) & " of " & Trim(Str(GameList.ListCount))
End Sub

Private Sub GameList_KeyPress(KeyAscii As Integer)
    If KeyAscii = Val(vbKeyEscape) Then ExitLauncher
    If KeyAscii = Val(vbKeyReturn) Then RunGame
End Sub

Sub ExitLauncher()
    Unload Me
    End
End Sub

Sub RunGame()
    On Error Resume Next
    RomList.ListIndex = GameList.ListIndex
    ChangeDir cfgNewEmulatorDir
    If GameList.ListIndex >= 0 Then Shell (Chr$(34) & cfgNewEmulatorDir & cfgEmulatorExe & Chr$(34) & " " & Chr$(34) & ProgramPath & RomsDir & "\" & RomList.FileName & Chr$(34)), vbNormalFocus
    On Error GoTo 0
End Sub

Sub CopyEmulator()
    If ObjFso.FileExists(cfgNewEmulatorDir & cfgEmulatorExe) = False Then
        If ObjFso.FileExists(ProgramPath & EmulatorDir & "\" & cfgEmulatorExe) = True Then
            For x = 1 To Len(cfgNewEmulatorDir)
                If Mid(cfgNewEmulatorDir, x, 1) = "\" Then
                    If ObjFso.FolderExists(Left(cfgNewEmulatorDir, x)) = False Then ObjFso.CreateFolder (Left(cfgNewEmulatorDir, x))
                End If
            Next x
            ObjFso.CopyFolder ProgramPath & EmulatorDir, Left(cfgNewEmulatorDir, Len(cfgNewEmulatorDir) - 1), True
            ResetAttributes (cfgNewEmulatorDir)
        Else
            MsgBox "Emulator: """ & ProgramPath & EmulatorDir & "\" & cfgEmulatorExe & """ not found!", vbCritical + vbOKOnly, "Error"
            ExitLauncher
        End If
    End If
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

Sub ReadConfig()
    Dim temp As String
    If ObjFso.FileExists(ProgramPath & ConfigFile) = True Then
        Open ProgramPath & ConfigFile For Input As #1
        Do
            Line Input #1, temp
            temp = Trim(temp)
            If InStr(temp, ";") <> 1 And InStr(temp, "=") > 0 Then
                If InStr(LCase(temp), "windowtitle") = 1 Then cfgWindowTitle = Trim(Mid(temp, (InStr(temp, "=") + 1), Len(temp)))
                If InStr(LCase(temp), "emulatorexe") = 1 Then cfgEmulatorExe = Trim(Mid(temp, (InStr(temp, "=") + 1), Len(temp)))
                If InStr(LCase(temp), "copyemulatorto") = 1 Then cfgNewEmulatorDir = Trim(Mid(temp, (InStr(temp, "=") + 1), Len(temp)))
                If InStr(LCase(temp), "romsext") = 1 Then cfgRomsExt = Trim(Mid(temp, (InStr(temp, "=") + 1), Len(temp)))
                If InStr(LCase(temp), "imagesext") = 1 Then cfgImagesExt = Trim(Mid(temp, (InStr(temp, "=") + 1), Len(temp)))
            End If
        Loop Until EOF(1)
        Close #1
    Else
        MsgBox "Config file: """ & ProgramPath & ConfigFile & """ not found!", vbCritical + vbOKOnly, "Error"
        ExitLauncher
    End If
End Sub

Sub ResetAttributes(strPath As String)
    Dim objFolder, objSubFolders, objFile, objSFldr

    Set objFolder = ObjFso.GetFolder(strPath)
    For Each objFile In objFolder.Files
        objFile.Attributes = 32
    Next

    Set objSubFolders = objFolder.SubFolders
    For Each objSFldr In objSubFolders
        ResetAttributes (objSFldr)
    Next

    Set objSubFolders = Nothing
    Set objFolder = Nothing
End Sub
