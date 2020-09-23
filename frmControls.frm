VERSION 5.00
Begin VB.Form frmControls 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Controls"
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2070
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   2070
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnOkay 
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
      Left            =   600
      TabIndex        =   5
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label lblGeneral 
      Caption         =   "Up: Power drop."
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   3135
   End
   Begin VB.Label lblGeneral 
      Caption         =   "Space: Rotate piece."
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   3135
   End
   Begin VB.Label lblGeneral 
      Caption         =   "Down: Drop piece lower."
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   3135
   End
   Begin VB.Label lblGeneral 
      Caption         =   "Right: Move piece right."
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   3135
   End
   Begin VB.Label lblGeneral 
      Caption         =   "Left: Move piece left."
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "frmControls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnOkay_Click()
    Unload Me
End Sub
