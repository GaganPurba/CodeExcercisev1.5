VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 1  'vbSimpleBound
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClScoreBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function StartMatch(ByVal HomeTeam As String, ByVal AwayTeam As String, ByVal StartTime As Date) As Boolean
    On Error Resume Next
    
    Dim xMatch As New ClMatchInfo
    xMatch.mHomeTeam = HomeTeam
    xMatch.mAwayTeam = AwayTeam
    xMatch.mStartTime = StartTime
        
    mCollScoreBoard.Add xMatch, HomeTeam & "||" & AwayTeam
    If Err.Number <> 0 Then
        MsgBox "This match already exists.", vbOKOnly
        Exit Function
    End If
    
    StartMatch = True
End Function

Public Function EndMatch(ByVal HomeTeam As String, ByVal HomeTeamScore As Integer, ByVal AwayTeam As String, ByVal AwayTeamScore As Integer, ByVal EndTime As Date) As Boolean
    On Error Resume Next
    
    Dim xMatch As ClMatchInfo
    Set xMatch = mCollScoreBoard(HomeTeam & "||" & AwayTeam)
    xMatch.mHomeTeamScore = HomeTeamScore
    xMatch.mAwayTeamScore = AwayTeamScore
    xMatch.mEndTime = EndTime
    
    EndMatch = True
End Function
