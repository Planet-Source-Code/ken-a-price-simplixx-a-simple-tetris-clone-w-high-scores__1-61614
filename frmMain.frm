VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Simplixx"
   ClientHeight    =   5625
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   4665
   ForeColor       =   &H00808080&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   375
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   311
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrFlash 
      Interval        =   100
      Left            =   3480
      Top             =   3840
   End
   Begin VB.PictureBox picNext 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   2880
      ScaleHeight     =   65
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   97
      TabIndex        =   9
      Top             =   480
      Width           =   1455
   End
   Begin VB.PictureBox picBlocks 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   1560
      Picture         =   "frmMain.frx":0E42
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   204
      TabIndex        =   0
      Top             =   5040
      Visible         =   0   'False
      Width           =   3060
   End
   Begin VB.PictureBox picGrid 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   5355
      Left            =   120
      ScaleHeight     =   357
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   170
      TabIndex        =   1
      Top             =   120
      Width           =   2550
      Begin VB.Label lblPaused 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Paused"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   495
         Left            =   240
         TabIndex        =   12
         Top             =   1080
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label lblNewgame 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "New Game Press F2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   855
         Left            =   360
         TabIndex        =   11
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label lblGameover 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Game Over"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   495
         Left            =   240
         TabIndex        =   10
         Top             =   1080
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label lblNewgame2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "New Game Press F2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   855
         Left            =   345
         TabIndex        =   13
         Top             =   855
         Width           =   1815
      End
   End
   Begin VB.Label lblLines 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   2760
      TabIndex        =   8
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label lblLevel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   2760
      TabIndex        =   7
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label Labels 
      BackStyle       =   0  'Transparent
      Caption         =   "Lines:"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   3
      Left            =   2760
      TabIndex        =   6
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label Labels 
      BackStyle       =   0  'Transparent
      Caption         =   "Level:"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   2
      Left            =   2760
      TabIndex        =   5
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Labels 
      BackStyle       =   0  'Transparent
      Caption         =   "Next:"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   1
      Left            =   2760
      TabIndex        =   4
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label lblScore 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   2760
      TabIndex        =   3
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Labels 
      BackStyle       =   0  'Transparent
      Caption         =   "Score:"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   0
      Left            =   2760
      TabIndex        =   2
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Line Lines 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   3
      Index           =   4
      X1              =   304
      X2              =   304
      Y1              =   367
      Y2              =   8
   End
   Begin VB.Line Lines 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   3
      Index           =   3
      X1              =   179
      X2              =   179
      Y1              =   367
      Y2              =   6
   End
   Begin VB.Line Lines 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   3
      Index           =   2
      X1              =   6
      X2              =   6
      Y1              =   367
      Y2              =   6
   End
   Begin VB.Line Lines 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   3
      Index           =   1
      X1              =   8
      X2              =   304
      Y1              =   6
      Y2              =   6
   End
   Begin VB.Line Lines 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   3
      Index           =   0
      X1              =   8
      X2              =   304
      Y1              =   367
      Y2              =   367
   End
   Begin VB.Menu mnuGame 
      Caption         =   "&Game"
      Begin VB.Menu mnuNewGame 
         Caption         =   "&New Game"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuPause 
         Caption         =   "&Pause"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuLine3431 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHighscores 
         Caption         =   "View Highscores"
      End
      Begin VB.Menu mnuLine3954 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "&Quit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuControls 
         Caption         =   "&Controls"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function PlaySound Lib _
    "winmm.dll" Alias "PlaySoundA" ( _
    ByVal lpszName As String, _
    ByVal hModule As Long, _
    ByVal dwFlags As Long) As Long

Private Const SND_ASYNC& = &H1
' Play asynchronously
Private Const SND_NODEFAULT& = &H2
' Silence if sound not found
Private Const SND_RESOURCE& = &H40004
' Name is resource name or atom

Dim hInst As Long
' Handle to Application Instance
Dim sSoundName As String
' String to hold sound resource name
Dim lFlags As Long
' PlaySound() flags
Dim lRet As Long
' Return value


    Dim CountBlocks, CountTime As Long, CountLines As Integer, LinesErased(3) As Integer
Dim LineAnimationColor As Integer
Dim LineAnimation As Boolean

Dim PieceAnimation As Integer
Dim CheatKeys As Integer
Dim CheatActivated As Boolean



Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long


Dim Grid(9, -2 To 20) As Boolean
Dim GridColor(9, -2 To 20) As Integer
Dim Level As Integer
Dim Speed(20) As Long 'speed of level
Dim NextPiece As Integer
Dim NextPieceGrid(3) As TetrisPiece
Dim Score As Long
Dim xLines As Integer
Dim Paused As Boolean
Dim ScoreDone As Boolean

Dim Gameover As Boolean 'Game Over


Private Sub NewGame()
    LineAnimation = False
    ScoreDone = False
    lblNewgame.Visible = False
    lblNewgame2.Visible = False
    lblGameover.Visible = False

    For a = 0 To 9
        For B = -2 To 20
            Grid(a, B) = False
        Next B
    Next a
    
    For a = 0 To 9
        For B = 0 To 20
            GridColor(a, B) = 0
        Next B
    Next a
    
    Level = 0
    Score = 0
    xLines = 0
    Gameover = False
    Paused = False
    lblLevel.Caption = 0
    UpdateStats
    
    Randomize
    NextPiece = Int(Rnd * 7)
    
    GameLoop
End Sub

Private Sub GameOver2()
    
    Dim msTime2 As Long
    Gameover = True
    
    For a = 20 To 0 Step -1
        msTime2 = 100 + GetTickCount
        
        For B = 0 To 9
            GridColor(B, a) = 7
        Next B
        DrawGrid
        
        
        sSoundName = "gameover"
        lFlags = SND_RESOURCE + SND_ASYNC + SND_NODEFAULT
        If ScoreDone = False Then lRet = PlaySound(sSoundName, hInst, lFlags)
            
        Do
            DoEvents
        Loop While GetTickCount < msTime2
    Next a
        
    lblGameover.Visible = True
        
    CurrentColor = 7
    
    
    Dim HighScoreName As String
    If ScoreDone = False And Score > Int(Val(ReadRegistry(HKEY_LOCAL_MACHINE, "Software/Simplixx", "Score1"))) Then
        ScoreDone = True
        HighScoreName = InputBox$("You have achieved a high score. Please enter your name.", "High Score", "Anonymous")
        
        WriteRegistry HKEY_LOCAL_MACHINE, "Software/Simplixx", "Score5", ValString, ReadRegistry(HKEY_LOCAL_MACHINE, "Software/Simplixx", "Score4")
        WriteRegistry HKEY_LOCAL_MACHINE, "Software/Simplixx", "Name5", ValString, ReadRegistry(HKEY_LOCAL_MACHINE, "Software/Simplixx", "Name4")
        WriteRegistry HKEY_LOCAL_MACHINE, "Software/Simplixx", "Score4", ValString, ReadRegistry(HKEY_LOCAL_MACHINE, "Software/Simplixx", "Score3")
        WriteRegistry HKEY_LOCAL_MACHINE, "Software/Simplixx", "Name4", ValString, ReadRegistry(HKEY_LOCAL_MACHINE, "Software/Simplixx", "Name3")
        WriteRegistry HKEY_LOCAL_MACHINE, "Software/Simplixx", "Score3", ValString, ReadRegistry(HKEY_LOCAL_MACHINE, "Software/Simplixx", "Score2")
        WriteRegistry HKEY_LOCAL_MACHINE, "Software/Simplixx", "Name3", ValString, ReadRegistry(HKEY_LOCAL_MACHINE, "Software/Simplixx", "Name2")
        WriteRegistry HKEY_LOCAL_MACHINE, "Software/Simplixx", "Score2", ValString, ReadRegistry(HKEY_LOCAL_MACHINE, "Software/Simplixx", "Score1")
        WriteRegistry HKEY_LOCAL_MACHINE, "Software/Simplixx", "Name2", ValString, ReadRegistry(HKEY_LOCAL_MACHINE, "Software/Simplixx", "Name1")
        
        WriteRegistry HKEY_LOCAL_MACHINE, "Software/Simplixx", "Score1", ValString, Score
        WriteRegistry HKEY_LOCAL_MACHINE, "Software/Simplixx", "Name1", ValString, HighScoreName
        frmHighscores.Show 1
        Exit Sub
        
    ElseIf ScoreDone = False And Score > Int(Val(ReadRegistry(HKEY_LOCAL_MACHINE, "Software/Simplixx", "Score2"))) Then
        ScoreDone = True
        HighScoreName = InputBox$("You have achieved a high score. Please enter your name.", "High Score", "Anonymous")
        
        WriteRegistry HKEY_LOCAL_MACHINE, "Software/Simplixx", "Score5", ValString, ReadRegistry(HKEY_LOCAL_MACHINE, "Software/Simplixx", "Score4")
        WriteRegistry HKEY_LOCAL_MACHINE, "Software/Simplixx", "Name5", ValString, ReadRegistry(HKEY_LOCAL_MACHINE, "Software/Simplixx", "Name4")
        WriteRegistry HKEY_LOCAL_MACHINE, "Software/Simplixx", "Score4", ValString, ReadRegistry(HKEY_LOCAL_MACHINE, "Software/Simplixx", "Score3")
        WriteRegistry HKEY_LOCAL_MACHINE, "Software/Simplixx", "Name4", ValString, ReadRegistry(HKEY_LOCAL_MACHINE, "Software/Simplixx", "Name3")
        WriteRegistry HKEY_LOCAL_MACHINE, "Software/Simplixx", "Score3", ValString, ReadRegistry(HKEY_LOCAL_MACHINE, "Software/Simplixx", "Score2")
        WriteRegistry HKEY_LOCAL_MACHINE, "Software/Simplixx", "Name3", ValString, ReadRegistry(HKEY_LOCAL_MACHINE, "Software/Simplixx", "Name2")
        
        WriteRegistry HKEY_LOCAL_MACHINE, "Software/Simplixx", "Score2", ValString, Score
        WriteRegistry HKEY_LOCAL_MACHINE, "Software/Simplixx", "Name2", ValString, HighScoreName
        frmHighscores.Show 1
        Exit Sub
        
    ElseIf ScoreDone = False And Score > Int(Val(ReadRegistry(HKEY_LOCAL_MACHINE, "Software/Simplixx", "Score3"))) Then
        ScoreDone = True
        HighScoreName = InputBox$("You have achieved a high score. Please enter your name.", "High Score", "Anonymous")
        
        WriteRegistry HKEY_LOCAL_MACHINE, "Software/Simplixx", "Score5", ValString, ReadRegistry(HKEY_LOCAL_MACHINE, "Software/Simplixx", "Score4")
        WriteRegistry HKEY_LOCAL_MACHINE, "Software/Simplixx", "Name5", ValString, ReadRegistry(HKEY_LOCAL_MACHINE, "Software/Simplixx", "Name4")
        WriteRegistry HKEY_LOCAL_MACHINE, "Software/Simplixx", "Score4", ValString, ReadRegistry(HKEY_LOCAL_MACHINE, "Software/Simplixx", "Score3")
        WriteRegistry HKEY_LOCAL_MACHINE, "Software/Simplixx", "Name4", ValString, ReadRegistry(HKEY_LOCAL_MACHINE, "Software/Simplixx", "Name3")
        
        WriteRegistry HKEY_LOCAL_MACHINE, "Software/Simplixx", "Score3", ValString, Score
        WriteRegistry HKEY_LOCAL_MACHINE, "Software/Simplixx", "Name3", ValString, HighScoreName
        frmHighscores.Show 1
        Exit Sub
        
    ElseIf ScoreDone = False And Score > Int(Val(ReadRegistry(HKEY_LOCAL_MACHINE, "Software/Simplixx", "Score4"))) Then
        ScoreDone = True
        HighScoreName = InputBox$("You have achieved a high score. Please enter your name.", "High Score", "Anonymous")
        
        WriteRegistry HKEY_LOCAL_MACHINE, "Software/Simplixx", "Score5", ValString, ReadRegistry(HKEY_LOCAL_MACHINE, "Software/Simplixx", "Score4")
        WriteRegistry HKEY_LOCAL_MACHINE, "Software/Simplixx", "Name5", ValString, ReadRegistry(HKEY_LOCAL_MACHINE, "Software/Simplixx", "Name4")
        
        WriteRegistry HKEY_LOCAL_MACHINE, "Software/Simplixx", "Score4", ValString, Score
        WriteRegistry HKEY_LOCAL_MACHINE, "Software/Simplixx", "Name4", ValString, HighScoreName
        frmHighscores.Show 1
        Exit Sub
        
    ElseIf ScoreDone = False And Score > Int(Val(ReadRegistry(HKEY_LOCAL_MACHINE, "Software/Simplixx", "Score5"))) Then
        ScoreDone = True
        HighScoreName = InputBox$("You have achieved a high score. Please enter your name.", "High Score", "Anonymous")
        
        WriteRegistry HKEY_LOCAL_MACHINE, "Software/Simplixx", "Score5", ValString, Score
        WriteRegistry HKEY_LOCAL_MACHINE, "Software/Simplixx", "Name5", ValString, HighScoreName
        frmHighscores.Show 1
        Exit Sub
    End If
    
    
    
End Sub

Private Sub CreatePiece()
    If Grid(3, 0) = True Or Grid(4, 0) = True Or Grid(5, 0) = True Or Grid(6, 0) = True Then
        GameOver2
        Exit Sub
    End If

    Randomize
    PieceName = NextPiece
    NextPiece = Int(Rnd * 7)
    Direction = 0
    
    Dim CheatPiece As Integer
    CheatPiece = Int(Rnd * 49)
    If CheatPiece = 0 Then NextPiece = 8
    
    Select Case NextPiece
        Case 0              '...OOOO...
                            '..........
            MakePiece(0).X = 3
            MakePiece(0).Y = 0
            MakePiece(1).X = 4
            MakePiece(1).Y = 0
            MakePiece(2).X = 5
            MakePiece(2).Y = 0
            MakePiece(3).X = 6
            MakePiece(3).Y = 0
            
        Case 1              '...OO.....
                            '....OO....
            MakePiece(0).X = 3
            MakePiece(0).Y = 0
            MakePiece(1).X = 4
            MakePiece(1).Y = 0
            MakePiece(2).X = 4
            MakePiece(2).Y = 1
            MakePiece(3).X = 5
            MakePiece(3).Y = 1
            
        Case 2              '....OO....
                            '...OO.....
            MakePiece(0).X = 3
            MakePiece(0).Y = 1
            MakePiece(1).X = 4
            MakePiece(1).Y = 0
            MakePiece(2).X = 4
            MakePiece(2).Y = 1
            MakePiece(3).X = 5
            MakePiece(3).Y = 0
            
        Case 3              '....O.....
                            '...OOO....
            MakePiece(0).X = 3
            MakePiece(0).Y = 1
            MakePiece(1).X = 4
            MakePiece(1).Y = 0
            MakePiece(2).X = 4
            MakePiece(2).Y = 1
            MakePiece(3).X = 5
            MakePiece(3).Y = 1
            
        Case 4              '...O......
                            '...OOO....
            MakePiece(0).X = 3
            MakePiece(0).Y = 0
            MakePiece(1).X = 3
            MakePiece(1).Y = 1
            MakePiece(2).X = 4
            MakePiece(2).Y = 1
            MakePiece(3).X = 5
            MakePiece(3).Y = 1
            
        Case 5              '.....O....
                            '...OOO....
            MakePiece(0).X = 3
            MakePiece(0).Y = 1
            MakePiece(1).X = 4
            MakePiece(1).Y = 1
            MakePiece(2).X = 5
            MakePiece(2).Y = 0
            MakePiece(3).X = 5
            MakePiece(3).Y = 1
            
        Case 6              '...OO.....
                            '...OO.....
            MakePiece(0).X = 3
            MakePiece(0).Y = 0
            MakePiece(1).X = 3
            MakePiece(1).Y = 1
            MakePiece(2).X = 4
            MakePiece(2).Y = 0
            MakePiece(3).X = 4
            MakePiece(3).Y = 1
            CurrentColor = 6
            
        Case 8
            MakePiece(0).X = 3
            MakePiece(0).Y = 0
            MakePiece(1).X = 3
            MakePiece(1).Y = 0
            MakePiece(2).X = 3
            MakePiece(2).Y = 0
            MakePiece(3).X = 3
            MakePiece(3).Y = 0
            CurrentColor = 8
    End Select
    
    For a = 0 To 3
        NextPieceGrid(a).X = MakePiece(a).X
        NextPieceGrid(a).Y = MakePiece(a).Y
    Next a
    
    Select Case PieceName
        Case 0              '...OOOO...
                            '..........
            MakePiece(0).X = 3
            MakePiece(0).Y = 0
            MakePiece(1).X = 4
            MakePiece(1).Y = 0
            MakePiece(2).X = 5
            MakePiece(2).Y = 0
            MakePiece(3).X = 6
            MakePiece(3).Y = 0
            CurrentColor = 0
            
        Case 1              '...OO.....
                            '....OO....
            MakePiece(0).X = 3
            MakePiece(0).Y = 0
            MakePiece(1).X = 4
            MakePiece(1).Y = 0
            MakePiece(2).X = 4
            MakePiece(2).Y = 1
            MakePiece(3).X = 5
            MakePiece(3).Y = 1
            CurrentColor = 1
            
        Case 2              '....OO....
                            '...OO.....
            MakePiece(0).X = 3
            MakePiece(0).Y = 1
            MakePiece(1).X = 4
            MakePiece(1).Y = 0
            MakePiece(2).X = 4
            MakePiece(2).Y = 1
            MakePiece(3).X = 5
            MakePiece(3).Y = 0
            CurrentColor = 2
            
        Case 3              '....O.....
                            '...OOO....
            MakePiece(0).X = 3
            MakePiece(0).Y = 1
            MakePiece(1).X = 4
            MakePiece(1).Y = 0
            MakePiece(2).X = 4
            MakePiece(2).Y = 1
            MakePiece(3).X = 5
            MakePiece(3).Y = 1
            CurrentColor = 3
            
        Case 4              '...O......
                            '...OOO....
            MakePiece(0).X = 3
            MakePiece(0).Y = 0
            MakePiece(1).X = 3
            MakePiece(1).Y = 1
            MakePiece(2).X = 4
            MakePiece(2).Y = 1
            MakePiece(3).X = 5
            MakePiece(3).Y = 1
            CurrentColor = 4
            
        Case 5              '.....O....
                            '...OOO....
            MakePiece(0).X = 3
            MakePiece(0).Y = 1
            MakePiece(1).X = 4
            MakePiece(1).Y = 1
            MakePiece(2).X = 5
            MakePiece(2).Y = 0
            MakePiece(3).X = 5
            MakePiece(3).Y = 1
            CurrentColor = 5
            
        Case 6              '...OO.....
                            '...OO.....
            MakePiece(0).X = 3
            MakePiece(0).Y = 0
            MakePiece(1).X = 3
            MakePiece(1).Y = 1
            MakePiece(2).X = 4
            MakePiece(2).Y = 0
            MakePiece(3).X = 4
            MakePiece(3).Y = 1
            CurrentColor = 6
            
        Case 8
            MakePiece(0).X = 3
            MakePiece(0).Y = 0
            MakePiece(1).X = 3
            MakePiece(1).Y = 0
            MakePiece(2).X = 3
            MakePiece(2).Y = 0
            MakePiece(3).X = 3
            MakePiece(3).Y = 0
            CurrentColor = 8
    End Select
    
    For a = 0 To 3
        CurrentPiece(a).X = MakePiece(a).X
        CurrentPiece(a).Y = MakePiece(a).Y
    Next a
    
    picNext.Cls
    For a = 0 To 3
        BitBlt picNext.hDC, (NextPieceGrid(a).X - 3) * 17 + 5, NextPieceGrid(a).Y * 17, 17, 17, picBlocks.hDC, NextPiece * 17, 0, vbSrcCopy
    Next a
End Sub

Private Sub LevelUp()
    
    If Int(xLines / 10) < 21 Then Level = Int(xLines / 10)
    
    lblLevel.Caption = Level
End Sub

Private Sub GameLoop()
    If Paused = True Then GoTo PauseSkip
    Dim msTime As Long
    Speed(0) = 1000
    Speed(1) = 900
    Speed(2) = 850
    Speed(3) = 800
    Speed(4) = 700
    Speed(5) = 650
    Speed(6) = 500
    Speed(7) = 450
    Speed(8) = 400
    Speed(9) = 400
    Speed(10) = 350
    Speed(11) = 350
    Speed(12) = 300
    Speed(13) = 310
    Speed(14) = 300
    Speed(15) = 270
    Speed(16) = 130
    Speed(17) = 100
    Speed(18) = 70
    Speed(19) = 50
    Speed(20) = 20
    
    

    CreatePiece

    
    DrawGrid
    
PauseSkip:
    Do
        Dim AppSpeed As Long
        Dim ColorAnimation As Boolean
        AppSpeed = Speed(Level)
        If PieceName = 8 Then AppSpeed = Int(AppSpeed / 4)
        msTime = AppSpeed + GetTickCount

        'Refresh grid
        DrawGrid
        
        'Check lines for level
        LevelUp
        
        Do While Paused = True
            lblPaused.Visible = True
                DoEvents
        Loop
        lblPaused.Visible = False
        
        Do
            DoEvents
        Loop While GetTickCount < msTime
        
        
        DropPiece 'Drop Piece
        
    Loop Until Gameover = True
End Sub

Public Sub DrawGrid()
'    If Gameover = True Then Exit Sub
    picGrid.Cls
    
    If Gameover = False Then
    For a = 0 To 3
        BitBlt picGrid.hDC, CurrentPiece(a).X * 17, CurrentPiece(a).Y * 17, 17, 17, picBlocks.hDC, CurrentColor * 17, 0, vbSrcCopy
    Next a
    End If

    For a = 0 To 9
        For B = 0 To 20
            If Grid(a, B) Then
                BitBlt picGrid.hDC, a * 17, B * 17, 17, 17, picBlocks.hDC, GridColor(a, B) * 17, 0, vbSrcCopy
            End If
        Next B
    Next a
End Sub



Private Sub Form_Load()
    MsgBox "COMPILE FIRST!!" & vbCrLf & "This version of Simplixx was released on pscode.com" & vbCrLf & "http://www.chaoticlogic.net/", vbOKOnly Or vbExclamation, "Simplixx"

    hInst = App.hInstance
    Paused = True
    Me.Show
    
    
    'NewGame
    
    Intro
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub Intro()
    For a = 0 To 9
        For B = 0 To 20
            Grid(a, B) = 1
            GridColor(a, B) = Int(Rnd * 6)
        Next B
    Next a
    
    DrawGrid
    
    For a = 0 To 9
        For B = 0 To 20
            Grid(a, B) = 0
            GridColor(a, B) = 0
        Next B
    Next a
End Sub

Private Sub UpdateStats()
    lblScore.Caption = Score
    lblLines.Caption = xLines
End Sub

Private Sub EndDrop()
    If Gameover = True Then Exit Sub
    sSoundName = "drop"
    lFlags = SND_RESOURCE + SND_ASYNC + SND_NODEFAULT
    lRet = PlaySound(sSoundName, hInst, lFlags)
    
    If PieceName = 8 Then
        UpdateStats
        GameLoop
    End If

    For a = 0 To 3
        Grid(CurrentPiece(a).X, CurrentPiece(a).Y) = True
        GridColor(CurrentPiece(a).X, CurrentPiece(a).Y) = CurrentColor
    Next a
    
    CountLines = 0
    CountTime = 0
    CountLines = 0
    
    For a = 0 To 20
        CountBlocks = 0
        
        For B = 0 To 9
            If Grid(B, a) = True Then CountBlocks = CountBlocks + 1
        Next B
        
        If CountBlocks = 10 Then
            CountLines = CountLines + 1
            LinesErased(CountLines - 1) = a
        End If
    Next a
    
    xLines = xLines + CountLines
    
    Select Case CountLines
        Case 1
            Score = Score + 300
            sSoundName = "line"
            lFlags = SND_RESOURCE + SND_ASYNC + SND_NODEFAULT
            lRet = PlaySound(sSoundName, hInst, lFlags)
        Case 2
            Score = Score + 600
            sSoundName = "line"
            lFlags = SND_RESOURCE + SND_ASYNC + SND_NODEFAULT
            lRet = PlaySound(sSoundName, hInst, lFlags)
        Case 3
            Score = Score + 1200
            sSoundName = "line"
            lFlags = SND_RESOURCE + SND_ASYNC + SND_NODEFAULT
            lRet = PlaySound(sSoundName, hInst, lFlags)
        Case 4
            Score = Score + 2500
            sSoundName = "tetris"
            lFlags = SND_RESOURCE + SND_ASYNC + SND_NODEFAULT
            lRet = PlaySound(sSoundName, hInst, lFlags)
    End Select
    
    'Remove Pieces
    
    If CountLines > 0 Then
        LineAnimationColor = 10
        For a = 0 To CountLines - 1
            For B = 0 To 9
                GridColor(B, LinesErased(a)) = LineAnimationColor
            Next B
        Next a
    
        DrawGrid
        LineAnimation = True
        
        Dim RemoveTime As Long
        RemoveTime = 500 + GetTickCount
        Do
            DoEvents
        Loop While GetTickCount < RemoveTime
        
        LineAnimation = False
    
        For a = 0 To CountLines - 1
            For B = 0 To 9
                Grid(B, LinesErased(a)) = False
                GridColor(B, LinesErased(a)) = 0
            Next B
            
            For B = 0 To 9
                For c = LinesErased(a) To 1 Step -1
                    Grid(B, c) = Grid(B, c - 1)
                    GridColor(B, c) = GridColor(B, c - 1)
                Next c
            Next B
        Next a
    End If
    
    UpdateStats
    
    GameLoop
End Sub

Private Sub PowerDrop()
    If PieceName = 8 Then
        For a = CurrentPiece(0).Y To 20
            Grid(CurrentPiece(0).X, a) = False
            DrawGrid
        Next a
        
        Score = Score + Abs(20 - CurrentPiece(0).Y) * 50
        EndDrop
        Exit Sub
    End If

    For a = 0 To 3
        If CurrentPiece(a).Y + 1 = 21 Then
            EndDrop
            Exit Sub
        ElseIf Grid(CurrentPiece(a).X, CurrentPiece(a).Y + 1) = True Then
            EndDrop
            Exit Sub
        End If
    Next a
    
    Dim DropLines As Integer
    DropLines = 0
    
    Do
        DropLines = DropLines + 1
        For a = 0 To 3
            CurrentPiece(a).Y = CurrentPiece(a).Y + 1
        Next a
        
        For a = 0 To 3
            If CurrentPiece(a).Y + 1 = 21 Then
                GoTo CountLines
            ElseIf Grid(CurrentPiece(a).X, CurrentPiece(a).Y + 1) = True Then
                GoTo CountLines
            End If
        Next a
    Loop
    
CountLines:
    Score = Score + (DropLines * 5)
    UpdateStats
    EndDrop
End Sub

Private Sub DropPiece()
    If PieceName = 8 Then
        If CurrentPiece(0).Y + 1 = 21 Then
            EndDrop
            Exit Sub
        ElseIf Grid(CurrentPiece(0).X, CurrentPiece(1).Y + 1) = True Then
            Grid(CurrentPiece(0).X, CurrentPiece(1).Y + 1) = False
            Score = Score + 100
            UpdateStats
        End If
        For a = 0 To 3
            CurrentPiece(a).Y = CurrentPiece(a).Y + 1
        Next a
        DrawGrid
        Exit Sub
    End If

    For a = 0 To 3
        If CurrentPiece(a).Y + 1 = 21 Then
            EndDrop
            Exit Sub
        ElseIf Grid(CurrentPiece(a).X, CurrentPiece(a).Y + 1) = True Then
            EndDrop
            Exit Sub
        End If
    Next a
    
    For a = 0 To 3
        CurrentPiece(a).Y = CurrentPiece(a).Y + 1
    Next a
    
    DrawGrid
    
End Sub

Private Sub MoveLeft()
    If PieceName = 8 Then
        If CurrentPiece(0).X - 1 = -1 Then
            Exit Sub
        ElseIf Grid(CurrentPiece(0).X - 1, CurrentPiece(0).Y) = True Then
            Score = Score + 100
            UpdateStats
            Grid(CurrentPiece(0).X - 1, CurrentPiece(0).Y) = False
        End If
        For a = 0 To 3
            CurrentPiece(a).X = CurrentPiece(a).X - 1
        Next a
        DrawGrid
        Exit Sub
    End If

    For a = 0 To 3
        If CurrentPiece(a).X - 1 = -1 Then
            Exit Sub
        ElseIf Grid(CurrentPiece(a).X - 1, CurrentPiece(a).Y) = True Then
            Exit Sub
        End If
    Next a

    For a = 0 To 3
        CurrentPiece(a).X = CurrentPiece(a).X - 1
    Next a
    
    DrawGrid
    
End Sub

Private Sub MoveRight()
    If PieceName = 8 Then
        If CurrentPiece(0).X + 1 = 10 Then
            Exit Sub
        ElseIf Grid(CurrentPiece(0).X + 1, CurrentPiece(0).Y) = True Then
            Score = Score + 100
            UpdateStats
        
            Grid(CurrentPiece(0).X + 1, CurrentPiece(0).Y) = False
        End If
        For a = 0 To 3
            CurrentPiece(a).X = CurrentPiece(a).X + 1
        Next a
        DrawGrid
        Exit Sub
    End If
    
    For a = 0 To 3
        If CurrentPiece(a).X + 1 = 10 Then
            Exit Sub
        ElseIf Grid(CurrentPiece(a).X + 1, CurrentPiece(a).Y) = True Then
            Exit Sub
        End If
    Next a

    For a = 0 To 3
        CurrentPiece(a).X = CurrentPiece(a).X + 1
    Next a
    
    DrawGrid
    
End Sub

Private Sub RotatePiece()
    Dim RotSource As TetrisPiece
    Select Case PieceName
        Case 0              '...OOOO...
                            '..........
            Select Case Direction
                Case 0
                    RotSource.X = CurrentPiece(0).X + 1
                    RotSource.Y = CurrentPiece(0).Y - 1
                    
                    'Check piece
                    If RotSource.Y + 1 > 20 Then Exit Sub
                    If RotSource.Y + 2 > 20 Then Exit Sub
                    If Grid(RotSource.X + 1, RotSource.Y) = False And _
                      Grid(RotSource.X + 1, RotSource.Y + 1) = False And _
                      Grid(RotSource.X + 1, RotSource.Y + 2) = False And _
                      Grid(RotSource.X + 1, RotSource.Y + 3) = False Then
                        'Rotate Piece
                        Direction = Direction + 1
                        CurrentPiece(0).X = RotSource.X + 1
                        CurrentPiece(0).Y = RotSource.Y - 1 + 1
                        CurrentPiece(1).X = RotSource.X + 1
                        CurrentPiece(1).Y = RotSource.Y + 1
                        CurrentPiece(2).X = RotSource.X + 1
                        CurrentPiece(2).Y = RotSource.Y + 1 + 1
                        CurrentPiece(3).X = RotSource.X + 1
                        CurrentPiece(3).Y = RotSource.Y + 2 + 1
                        
                        'Refresh
                        DrawGrid
                        Exit Sub
                    End If
                Case 1
                    RotSource.X = CurrentPiece(0).X - 1
                    RotSource.Y = CurrentPiece(0).Y + 1
                    
                    'Check piece
                    If RotSource.X - 1 < 0 Then Exit Sub
                    If RotSource.X + 2 > 9 Then Exit Sub
                    If RotSource.X + 1 > 9 Then Exit Sub
                    If Grid(RotSource.X - 1, RotSource.Y) = False And _
                      Grid(RotSource.X, RotSource.Y) = False And _
                      Grid(RotSource.X + 1, RotSource.Y) = False And _
                      Grid(RotSource.X + 2, RotSource.Y) = False Then
                        'Rotate Piece
                        Direction = 0
                        CurrentPiece(0).X = RotSource.X - 1
                        CurrentPiece(0).Y = RotSource.Y - 1 + 1
                        CurrentPiece(1).X = RotSource.X
                        CurrentPiece(1).Y = RotSource.Y - 1 + 1
                        CurrentPiece(2).X = RotSource.X + 1
                        CurrentPiece(2).Y = RotSource.Y - 1 + 1
                        CurrentPiece(3).X = RotSource.X + 2
                        CurrentPiece(3).Y = RotSource.Y - 1 + 1
                        
                        'Refresh
                        DrawGrid
                        Exit Sub
                    End If
            End Select
            
        Case 1              '...OO.....
                            '....OO....
            Select Case Direction
                Case 0
                    RotSource.X = CurrentPiece(0).X + 1
                    RotSource.Y = CurrentPiece(0).Y + 1
                    
                    'Check piece
                    If RotSource.X + 1 > 9 Then Exit Sub
                    If RotSource.Y + 1 > 20 Then Exit Sub
                    If Grid(RotSource.X + 1, RotSource.Y - 1) = False And _
                      Grid(RotSource.X + 1, RotSource.Y) = False And _
                      Grid(RotSource.X, RotSource.Y) = False And _
                      Grid(RotSource.X, RotSource.Y + 1) = False Then
                        'Rotate Piece
                        Direction = Direction + 1
                        CurrentPiece(0).X = RotSource.X + 1
                        CurrentPiece(0).Y = RotSource.Y - 1
                        CurrentPiece(1).X = RotSource.X + 1
                        CurrentPiece(1).Y = RotSource.Y
                        CurrentPiece(2).X = RotSource.X
                        CurrentPiece(2).Y = RotSource.Y
                        CurrentPiece(3).X = RotSource.X
                        CurrentPiece(3).Y = RotSource.Y + 1
                        
                        'Refresh
                        DrawGrid
                        Exit Sub
                    End If
                Case 1
                    RotSource.X = CurrentPiece(0).X - 1
                    RotSource.Y = CurrentPiece(0).Y + 1
                    
                    'Check piece
                    If RotSource.X - 1 < 0 Then Exit Sub
                    If RotSource.Y + 1 > 20 Then Exit Sub
                    If Grid(RotSource.X - 1, RotSource.Y - 1) = False And _
                      Grid(RotSource.X, RotSource.Y - 1) = False And _
                      Grid(RotSource.X, RotSource.Y) = False And _
                      Grid(RotSource.X + 1, RotSource.Y) = False Then
                        'Rotate Piece
                        Direction = 0
                        CurrentPiece(0).X = RotSource.X - 1
                        CurrentPiece(0).Y = RotSource.Y - 1
                        CurrentPiece(1).X = RotSource.X
                        CurrentPiece(1).Y = RotSource.Y - 1
                        CurrentPiece(2).X = RotSource.X
                        CurrentPiece(2).Y = RotSource.Y
                        CurrentPiece(3).X = RotSource.X + 1
                        CurrentPiece(3).Y = RotSource.Y
                        
                        'Refresh
                        DrawGrid
                        Exit Sub
                    End If
                    
            End Select
        Case 2              '....OO....
                            '...OO.....
            Select Case Direction
                Case 0
                    RotSource.X = CurrentPiece(0).X + 1
                    RotSource.Y = CurrentPiece(0).Y
            
                    'Check piece
                    If RotSource.Y + 1 > 20 Then Exit Sub
                    If Grid(RotSource.X, RotSource.Y - 1) = False And _
                      Grid(RotSource.X, RotSource.Y) = False And _
                      Grid(RotSource.X + 1, RotSource.Y) = False And _
                      Grid(RotSource.X + 1, RotSource.Y + 1) = False Then
                        'Rotate piece
                        Direction = Direction + 1
                        CurrentPiece(0).X = RotSource.X
                        CurrentPiece(0).Y = RotSource.Y - 1
                        CurrentPiece(1).X = RotSource.X
                        CurrentPiece(1).Y = RotSource.Y
                        CurrentPiece(2).X = RotSource.X + 1
                        CurrentPiece(2).Y = RotSource.Y
                        CurrentPiece(3).X = RotSource.X + 1
                        CurrentPiece(3).Y = RotSource.Y + 1
                        
                        'Refresh
                        DrawGrid
                        Exit Sub
                    End If
                Case 1
                    RotSource.X = CurrentPiece(0).X
                    RotSource.Y = CurrentPiece(0).Y + 1
            
                    'Check piece
                    If RotSource.X - 1 < 0 Then Exit Sub
                    If Grid(RotSource.X - 1, RotSource.Y) = False And _
                      Grid(RotSource.X, RotSource.Y) = False And _
                      Grid(RotSource.X, RotSource.Y - 1) = False And _
                      Grid(RotSource.X + 1, RotSource.Y - 1) = False Then
                        'Rotate piece
                        Direction = 0
                        CurrentPiece(0).X = RotSource.X - 1
                        CurrentPiece(0).Y = RotSource.Y
                        CurrentPiece(1).X = RotSource.X
                        CurrentPiece(1).Y = RotSource.Y
                        CurrentPiece(2).X = RotSource.X
                        CurrentPiece(2).Y = RotSource.Y - 1
                        CurrentPiece(3).X = RotSource.X + 1
                        CurrentPiece(3).Y = RotSource.Y - 1
                        
                        'Refresh
                        DrawGrid
                        Exit Sub
                    End If
            End Select
        Case 3              '....O.....
                            '...OOO....
            Select Case Direction
                Case 0
                    RotSource.X = CurrentPiece(2).X
                    RotSource.Y = CurrentPiece(2).Y
            
                    'Check piece
                    If RotSource.Y + 1 > 20 Then Exit Sub
                    If Grid(RotSource.X, RotSource.Y + 1) = False Then
                        'Rotate piece
                        Direction = Direction + 1
                        CurrentPiece(0).X = RotSource.X
                        CurrentPiece(0).Y = RotSource.Y - 1
                        CurrentPiece(1).X = RotSource.X + 1
                        CurrentPiece(1).Y = RotSource.Y
                        CurrentPiece(2).X = RotSource.X
                        CurrentPiece(2).Y = RotSource.Y
                        CurrentPiece(3).X = RotSource.X
                        CurrentPiece(3).Y = RotSource.Y + 1
                        
                        'Refresh
                        DrawGrid
                        Exit Sub
                    End If
                Case 1
                    RotSource.X = CurrentPiece(2).X
                    RotSource.Y = CurrentPiece(2).Y
            
                    'Check piece
                    If RotSource.X - 1 < 0 Then Exit Sub
                    If Grid(RotSource.X - 1, RotSource.Y) = False Then
                        'Rotate piece
                        Direction = Direction + 1
                        CurrentPiece(0).X = RotSource.X + 1
                        CurrentPiece(0).Y = RotSource.Y
                        CurrentPiece(1).X = RotSource.X
                        CurrentPiece(1).Y = RotSource.Y + 1
                        CurrentPiece(2).X = RotSource.X
                        CurrentPiece(2).Y = RotSource.Y
                        CurrentPiece(3).X = RotSource.X - 1
                        CurrentPiece(3).Y = RotSource.Y
                        
                        'Refresh
                        DrawGrid
                        Exit Sub
                    End If
                Case 2
                    RotSource.X = CurrentPiece(2).X
                    RotSource.Y = CurrentPiece(2).Y
            
                    'Check piece
                    If RotSource.Y + 1 > 20 Then Exit Sub
                    If Grid(RotSource.X, RotSource.Y + 1) = False Then
                        'Rotate piece
                        Direction = Direction + 1
                        CurrentPiece(0).X = RotSource.X
                        CurrentPiece(0).Y = RotSource.Y - 1
                        CurrentPiece(1).X = RotSource.X - 1
                        CurrentPiece(1).Y = RotSource.Y
                        CurrentPiece(2).X = RotSource.X
                        CurrentPiece(2).Y = RotSource.Y
                        CurrentPiece(3).X = RotSource.X
                        CurrentPiece(3).Y = RotSource.Y + 1
                        
                        'Refresh
                        DrawGrid
                        Exit Sub
                    End If
                Case 3
                    RotSource.X = CurrentPiece(2).X
                    RotSource.Y = CurrentPiece(2).Y
            
                    'Check piece
                    If RotSource.X + 1 > 9 Then Exit Sub
                    If Grid(RotSource.X + 1, RotSource.Y) = False Then
                        'Rotate piece
                        Direction = 0
                        CurrentPiece(0).X = RotSource.X + 1
                        CurrentPiece(0).Y = RotSource.Y
                        CurrentPiece(1).X = RotSource.X
                        CurrentPiece(1).Y = RotSource.Y - 1
                        CurrentPiece(2).X = RotSource.X
                        CurrentPiece(2).Y = RotSource.Y
                        CurrentPiece(3).X = RotSource.X - 1
                        CurrentPiece(3).Y = RotSource.Y
                        
                        'Refresh
                        DrawGrid
                        Exit Sub
                    End If
            End Select
        Case 4              '...O......
                            '...OOO....
            Select Case Direction
                Case 0
                    RotSource.X = CurrentPiece(0).X + 1
                    RotSource.Y = CurrentPiece(0).Y + 1
            
                    'Check piece
                    If RotSource.Y + 1 > 20 Then Exit Sub
                    If Grid(RotSource.X, RotSource.Y - 1) = False And _
                      Grid(RotSource.X + 1, RotSource.Y - 1) = False And _
                      Grid(RotSource.X, RotSource.Y) = False And _
                      Grid(RotSource.X, RotSource.Y + 1) = False Then
                        'Rotate piece
                        Direction = Direction + 1
                        CurrentPiece(0).X = RotSource.X
                        CurrentPiece(0).Y = RotSource.Y - 1
                        CurrentPiece(1).X = RotSource.X + 1
                        CurrentPiece(1).Y = RotSource.Y - 1
                        CurrentPiece(2).X = RotSource.X
                        CurrentPiece(2).Y = RotSource.Y
                        CurrentPiece(3).X = RotSource.X
                        CurrentPiece(3).Y = RotSource.Y + 1
                        
                        'Refresh
                        DrawGrid
                        Exit Sub
                    End If
                Case 1
                    RotSource.X = CurrentPiece(0).X
                    RotSource.Y = CurrentPiece(0).Y + 1
            
                    'Check piece
                    If RotSource.Y + 1 > 20 Then Exit Sub
                    If RotSource.X - 1 < 0 Then Exit Sub
                    If Grid(RotSource.X - 1, RotSource.Y) = False And _
                      Grid(RotSource.X, RotSource.Y) = False And _
                      Grid(RotSource.X + 1, RotSource.Y) = False And _
                      Grid(RotSource.X + 1, RotSource.Y + 1) = False Then
                        'Rotate piece
                        Direction = Direction + 1
                        CurrentPiece(0).X = RotSource.X - 1
                        CurrentPiece(0).Y = RotSource.Y
                        CurrentPiece(1).X = RotSource.X
                        CurrentPiece(1).Y = RotSource.Y
                        CurrentPiece(2).X = RotSource.X + 1
                        CurrentPiece(2).Y = RotSource.Y
                        CurrentPiece(3).X = RotSource.X + 1
                        CurrentPiece(3).Y = RotSource.Y + 1
                        
                        'Refresh
                        DrawGrid
                        Exit Sub
                    End If
                Case 2
                    RotSource.X = CurrentPiece(0).X + 1
                    RotSource.Y = CurrentPiece(0).Y
            
                    'Check piece
                    'If RotSource.Y + 1 > 20 Then Exit Sub
                    If Grid(RotSource.X - 1, RotSource.Y + 1) = False And _
                      Grid(RotSource.X, RotSource.Y + 1) = False And _
                      Grid(RotSource.X, RotSource.Y) = False And _
                      Grid(RotSource.X, RotSource.Y - 1) = False Then
                        'Rotate piece
                        Direction = Direction + 1
                        CurrentPiece(0).X = RotSource.X - 1
                        CurrentPiece(0).Y = RotSource.Y + 1
                        CurrentPiece(1).X = RotSource.X
                        CurrentPiece(1).Y = RotSource.Y + 1
                        CurrentPiece(2).X = RotSource.X
                        CurrentPiece(2).Y = RotSource.Y
                        CurrentPiece(3).X = RotSource.X
                        CurrentPiece(3).Y = RotSource.Y - 1
                        
                        'Refresh
                        DrawGrid
                        Exit Sub
                    End If
                Case 3
                    RotSource.X = CurrentPiece(0).X + 1
                    RotSource.Y = CurrentPiece(0).Y - 1
            
                    'Check piece
                    If RotSource.X + 1 > 9 Then Exit Sub
                    If Grid(RotSource.X - 1, RotSource.Y - 1) = False And _
                      Grid(RotSource.X - 1, RotSource.Y) = False And _
                      Grid(RotSource.X, RotSource.Y) = False And _
                      Grid(RotSource.X + 1, RotSource.Y) = False Then
                        'Rotate piece
                        Direction = 0
                        CurrentPiece(0).X = RotSource.X - 1
                        CurrentPiece(0).Y = RotSource.Y - 1
                        CurrentPiece(1).X = RotSource.X - 1
                        CurrentPiece(1).Y = RotSource.Y
                        CurrentPiece(2).X = RotSource.X
                        CurrentPiece(2).Y = RotSource.Y
                        CurrentPiece(3).X = RotSource.X + 1
                        CurrentPiece(3).Y = RotSource.Y
                        
                        'Refresh
                        DrawGrid
                        Exit Sub
                    End If
            End Select
        Case 5              '.....O....
                            '...OOO....
            Select Case Direction
                Case 0
                    RotSource.X = CurrentPiece(1).X
                    RotSource.Y = CurrentPiece(1).Y
            
                    'Check piece
                    If RotSource.Y + 1 > 20 Then Exit Sub
                    If Grid(RotSource.X, RotSource.Y - 1) = False And _
                      Grid(RotSource.X, RotSource.Y) = False And _
                      Grid(RotSource.X, RotSource.Y + 1) = False And _
                      Grid(RotSource.X + 1, RotSource.Y + 1) = False Then
                        'Rotate piece
                        Direction = Direction + 1
                        CurrentPiece(0).X = RotSource.X
                        CurrentPiece(0).Y = RotSource.Y - 1
                        CurrentPiece(1).X = RotSource.X
                        CurrentPiece(1).Y = RotSource.Y
                        CurrentPiece(2).X = RotSource.X
                        CurrentPiece(2).Y = RotSource.Y + 1
                        CurrentPiece(3).X = RotSource.X + 1
                        CurrentPiece(3).Y = RotSource.Y + 1
                        
                        'Refresh
                        DrawGrid
                        Exit Sub
                    End If
                Case 1
                    RotSource.X = CurrentPiece(1).X
                    RotSource.Y = CurrentPiece(1).Y
            
                    'Check piece
                    If RotSource.Y + 1 > 20 Then Exit Sub
                    If RotSource.X - 1 < 0 Then Exit Sub
                    If Grid(RotSource.X + 1, RotSource.Y) = False And _
                      Grid(RotSource.X, RotSource.Y) = False And _
                      Grid(RotSource.X - 1, RotSource.Y) = False And _
                      Grid(RotSource.X - 1, RotSource.Y + 1) = False Then
                        'Rotate piece
                        Direction = Direction + 1
                        CurrentPiece(0).X = RotSource.X + 1
                        CurrentPiece(0).Y = RotSource.Y
                        CurrentPiece(1).X = RotSource.X
                        CurrentPiece(1).Y = RotSource.Y
                        CurrentPiece(2).X = RotSource.X - 1
                        CurrentPiece(2).Y = RotSource.Y
                        CurrentPiece(3).X = RotSource.X - 1
                        CurrentPiece(3).Y = RotSource.Y + 1
                        
                        'Refresh
                        DrawGrid
                        Exit Sub
                    End If
                Case 2
                    RotSource.X = CurrentPiece(1).X
                    RotSource.Y = CurrentPiece(1).Y
            
                    'Check piece
                    If RotSource.Y + 1 > 20 Then Exit Sub
                    If Grid(RotSource.X, RotSource.Y + 1) = False And _
                      Grid(RotSource.X, RotSource.Y) = False And _
                      Grid(RotSource.X, RotSource.Y - 1) = False And _
                      Grid(RotSource.X - 1, RotSource.Y - 1) = False Then
                        'Rotate piece
                        Direction = Direction + 1
                        CurrentPiece(0).X = RotSource.X
                        CurrentPiece(0).Y = RotSource.Y + 1
                        CurrentPiece(1).X = RotSource.X
                        CurrentPiece(1).Y = RotSource.Y
                        CurrentPiece(2).X = RotSource.X
                        CurrentPiece(2).Y = RotSource.Y - 1
                        CurrentPiece(3).X = RotSource.X - 1
                        CurrentPiece(3).Y = RotSource.Y - 1
                        
                        'Refresh
                        DrawGrid
                        Exit Sub
                    End If
                Case 3
                    RotSource.X = CurrentPiece(1).X
                    RotSource.Y = CurrentPiece(1).Y
            
                    'Check piece
                    If RotSource.Y + 1 > 20 Then Exit Sub
                    If RotSource.X + 1 > 9 Then Exit Sub
                    If Grid(RotSource.X - 1, RotSource.Y) = False And _
                      Grid(RotSource.X, RotSource.Y) = False And _
                      Grid(RotSource.X + 1, RotSource.Y) = False And _
                      Grid(RotSource.X + 1, RotSource.Y - 1) = False Then
                        'Rotate piece
                        Direction = 0
                        CurrentPiece(0).X = RotSource.X - 1
                        CurrentPiece(0).Y = RotSource.Y
                        CurrentPiece(1).X = RotSource.X
                        CurrentPiece(1).Y = RotSource.Y
                        CurrentPiece(2).X = RotSource.X + 1
                        CurrentPiece(2).Y = RotSource.Y
                        CurrentPiece(3).X = RotSource.X + 1
                        CurrentPiece(3).Y = RotSource.Y - 1
                        
                        'Refresh
                        DrawGrid
                        Exit Sub
                    End If
            End Select
    End Select
End Sub


Private Sub lblLines_Click()

End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show 1
End Sub

Private Sub mnuControls_Click()
    frmControls.Show 1
End Sub

Private Sub mnuHighscores_Click()
    frmHighscores.Show 1
End Sub

Private Sub mnuNewGame_Click()
    NewGame
End Sub

Private Sub mnuPause_Click()
    If Gameover = True Then Exit Sub
    
    If Paused = True Then
        Paused = False
        Exit Sub
    Else
        Paused = True
        GameLoop
        Exit Sub
    End If
End Sub

Private Sub mnuQuit_Click()
    Unload Me
    End
End Sub

Private Sub picGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    If LineAnimation = True Then Exit Sub
    If KeyCode = vbKeyK Then CheatKeys = 1
    If KeyCode = vbKeyE And CheatKeys = 1 Then CheatKeys = 2
    If KeyCode = vbKeyN And CheatKeys = 2 Then CheatKeys = 3
    If KeyCode = vbKeyI And CheatKeys = 3 Then CheatKeys = 4
    If KeyCode = vbKeyS And CheatKeys = 4 Then CheatKeys = 5
    If KeyCode = vbKeyH And CheatKeys = 5 Then CheatKeys = 6
    If KeyCode = vbKeyO And CheatKeys = 6 Then CheatKeys = 7
    If KeyCode = vbKeyT And CheatKeys = 7 Then
        CheatKeys = 0
        If CheatActivated = False Then
            CheatActivated = True
            Exit Sub
        End If
        If CheatActivated = True Then
            CheatActivated = False
            Exit Sub
        End If
        
    End If
    


    If Paused = True Then Exit Sub
    If Gameover = True Then Exit Sub
    Select Case KeyCode
        Case 37 'Left
            MoveLeft
        Case 38 'Up
            PowerDrop
        Case 39 'Right
            MoveRight
        Case 40 'Down
            DropPiece
        Case 32 'Space
            RotatePiece
    End Select
End Sub

Private Sub tmrFlash_Timer()
    If LineAnimation = True Then
        For a = 0 To CountLines - 1
            For B = 0 To 9
                GridColor(B, LinesErased(a)) = LineAnimationColor
            Next B
        Next a
        
        DrawGrid
        
        If LineAnimationColor = 11 Then
            LineAnimationColor = 10
            Exit Sub
        Else:
            LineAnimationColor = 11
            Exit Sub
        End If
    End If
End Sub
