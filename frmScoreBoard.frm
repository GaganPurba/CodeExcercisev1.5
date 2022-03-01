VERSION 5.00
Begin VB.Form frmScoreBoard 
   Caption         =   "Score board"
   ClientHeight    =   5070
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7800
   LinkTopic       =   "Form1"
   ScaleHeight     =   5070
   ScaleWidth      =   7800
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lblScoreBoard 
      Height          =   4455
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   7095
   End
End
Attribute VB_Name = "frmScoreBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    lblScoreBoard.Caption = GetScoreBoard()
    Me.Show
End Sub

Private Sub Form_Terminate()
    Debug.Print "frmScoreBoard.Form_Terminate"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Me.Hide
    Unload Me
End Sub
