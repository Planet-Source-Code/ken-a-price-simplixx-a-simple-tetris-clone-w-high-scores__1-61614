VERSION 5.00
Begin VB.Form frmHighscores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "High Scores"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4080
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   4080
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   0
      Top             =   2280
      Width           =   855
   End
   Begin VB.CommandButton btnReset 
      Caption         =   "Reset"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label lblScore 
      Caption         =   "0"
      Height          =   255
      Index           =   4
      Left            =   2640
      TabIndex        =   13
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label lblScore 
      Caption         =   "0"
      Height          =   255
      Index           =   3
      Left            =   2640
      TabIndex        =   12
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label lblScore 
      Caption         =   "0"
      Height          =   255
      Index           =   2
      Left            =   2640
      TabIndex        =   11
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label lblScore 
      Caption         =   "0"
      Height          =   255
      Index           =   1
      Left            =   2640
      TabIndex        =   10
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label lblScore 
      Caption         =   "0"
      Height          =   255
      Index           =   0
      Left            =   2640
      TabIndex        =   9
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label lblNames 
      Caption         =   "Name"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   8
      Top             =   1920
      Width           =   2415
   End
   Begin VB.Label lblNames 
      Caption         =   "Name"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   2415
   End
   Begin VB.Label lblNames 
      Caption         =   "Name"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Label lblNames 
      Caption         =   "Name"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   2415
   End
   Begin VB.Label lblNames 
      Caption         =   "Name"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label lblHighscores 
      Caption         =   "Score:"
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
      Index           =   0
      Left            =   2640
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lblHighscores 
      Caption         =   "Name:"
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
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmHighscores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnOk_Click()
    Unload Me
End Sub

Private Sub btnReset_Click()
    Dim qReset As VbMsgBoxResult
    qReset = MsgBox("Are you sure you want to reset the high scores?", vbYesNoCancel, "Reset")
    If qReset = vbYes Then
        WriteRegistry HKEY_LOCAL_MACHINE, "Software/Simplixx", "Score1", ValString, 0
        WriteRegistry HKEY_LOCAL_MACHINE, "Software/Simplixx", "Score2", ValString, 0
        WriteRegistry HKEY_LOCAL_MACHINE, "Software/Simplixx", "Score3", ValString, 0
        WriteRegistry HKEY_LOCAL_MACHINE, "Software/Simplixx", "Score4", ValString, 0
        WriteRegistry HKEY_LOCAL_MACHINE, "Software/Simplixx", "Score5", ValString, 0
        WriteRegistry HKEY_LOCAL_MACHINE, "Software/Simplixx", "Name1", ValString, "Anonymous"
        WriteRegistry HKEY_LOCAL_MACHINE, "Software/Simplixx", "Name2", ValString, "Anonymous"
        WriteRegistry HKEY_LOCAL_MACHINE, "Software/Simplixx", "Name3", ValString, "Anonymous"
        WriteRegistry HKEY_LOCAL_MACHINE, "Software/Simplixx", "Name4", ValString, "Anonymous"
        WriteRegistry HKEY_LOCAL_MACHINE, "Software/Simplixx", "Name5", ValString, "Anonymous"
    End If
    
    lblNames(0).Caption = ReadRegistry(HKEY_LOCAL_MACHINE, "Software/Simplixx", "Name1")
    lblNames(1).Caption = ReadRegistry(HKEY_LOCAL_MACHINE, "Software/Simplixx", "Name2")
    lblNames(2).Caption = ReadRegistry(HKEY_LOCAL_MACHINE, "Software/Simplixx", "Name3")
    lblNames(3).Caption = ReadRegistry(HKEY_LOCAL_MACHINE, "Software/Simplixx", "Name4")
    lblNames(4).Caption = ReadRegistry(HKEY_LOCAL_MACHINE, "Software/Simplixx", "Name5")

    lblScore(0).Caption = ReadRegistry(HKEY_LOCAL_MACHINE, "Software/Simplixx", "Score1")
    lblScore(1).Caption = ReadRegistry(HKEY_LOCAL_MACHINE, "Software/Simplixx", "Score2")
    lblScore(2).Caption = ReadRegistry(HKEY_LOCAL_MACHINE, "Software/Simplixx", "Score3")
    lblScore(3).Caption = ReadRegistry(HKEY_LOCAL_MACHINE, "Software/Simplixx", "Score4")
    lblScore(4).Caption = ReadRegistry(HKEY_LOCAL_MACHINE, "Software/Simplixx", "Score5")
End Sub

Private Sub Form_Load()
    lblNames(0).Caption = ReadRegistry(HKEY_LOCAL_MACHINE, "Software/Simplixx", "Name1")
    lblNames(1).Caption = ReadRegistry(HKEY_LOCAL_MACHINE, "Software/Simplixx", "Name2")
    lblNames(2).Caption = ReadRegistry(HKEY_LOCAL_MACHINE, "Software/Simplixx", "Name3")
    lblNames(3).Caption = ReadRegistry(HKEY_LOCAL_MACHINE, "Software/Simplixx", "Name4")
    lblNames(4).Caption = ReadRegistry(HKEY_LOCAL_MACHINE, "Software/Simplixx", "Name5")

    lblScore(0).Caption = ReadRegistry(HKEY_LOCAL_MACHINE, "Software/Simplixx", "Score1")
    lblScore(1).Caption = ReadRegistry(HKEY_LOCAL_MACHINE, "Software/Simplixx", "Score2")
    lblScore(2).Caption = ReadRegistry(HKEY_LOCAL_MACHINE, "Software/Simplixx", "Score3")
    lblScore(3).Caption = ReadRegistry(HKEY_LOCAL_MACHINE, "Software/Simplixx", "Score4")
    lblScore(4).Caption = ReadRegistry(HKEY_LOCAL_MACHINE, "Software/Simplixx", "Score5")
End Sub
