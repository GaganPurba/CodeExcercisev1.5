VERSION 5.00
Begin VB.Form frmGame 
   Caption         =   "Game Form"
   ClientHeight    =   4305
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   4305
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEndGame 
      Caption         =   "End game"
      Height          =   495
      Left            =   480
      TabIndex        =   9
      Top             =   3480
      Width           =   3015
   End
   Begin VB.TextBox txtAwayTeamScore 
      Height          =   375
      Left            =   1440
      TabIndex        =   8
      Top             =   2880
      Width           =   1935
   End
   Begin VB.TextBox txtHomeTeamScore 
      Height          =   375
      Left            =   1440
      TabIndex        =   6
      Top             =   2280
      Width           =   1935
   End
   Begin VB.CommandButton cmdStartGame 
      Caption         =   "Start game"
      Height          =   495
      Left            =   480
      TabIndex        =   4
      Top             =   1560
      Width           =   3015
   End
   Begin VB.TextBox txtAwayTeam 
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   840
      Width           =   1935
   End
   Begin VB.TextBox txtHomeTeam 
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label lblAwayTeamScore 
      Caption         =   "Away team score:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label lblHomeTeamScore 
      Caption         =   "Home team score:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label lblAwayTeam 
      Caption         =   "Away team:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label lblHomeTeam 
      Caption         =   "Home team:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdEndGame_Click()
    If txtHomeTeam.Text = "" Or txtAwayTeam.Text = "" Then
        MsgBox "Please provide Team names", vbOKOnly
        Exit Sub
    End If
    
    If Not (IsNumeric(txtHomeTeamScore.Text) Or IsNumeric(txtAwayTeamScore.Text)) Then
        MsgBox "Please provide Team Scores as numeric values", vbOKOnly
        Exit Sub
    End If
    
    If mScoreBoard.EndMatch(txtHomeTeam.Text, txtHomeTeamScore.Text, txtAwayTeam.Text, txtAwayTeamScore.Text, Now) Then
        Me.Hide
    End If
End Sub

Private Sub cmdStartGame_Click()
    If txtHomeTeam.Text = "" Or txtAwayTeam.Text = "" Then
        MsgBox "Please provide Team names", vbOKOnly
        Exit Sub
    End If
    
    cmdEndGame.Enabled = mScoreBoard.StartMatch(txtHomeTeam.Text, txtAwayTeam.Text, Now)
End Sub

Private Sub Form_Load()
    Me.Show
End Sub

Private Sub Form_Terminate()
    Debug.Print "frmGame.Form_Terminate"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Me.Hide
    Unload Me
End Sub

