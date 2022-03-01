Attribute VB_Name = "modExcercise"
Option Explicit

Public mCollScoreBoard As Collection
Public mScoreBoard As ClScoreBoard

Private Function SortScoreBoard() As Collection
    '
End Function

Public Function GetScoreBoard() As String
    If mCollScoreBoard.Count = 0 Then Exit Function
    
    Dim lngCt As Long
    Dim xMatch As ClMatchInfo
    Dim sStr As String
    For lngCt = 1 To mCollScoreBoard.Count
        Set xMatch = mCollScoreBoard(lngCt)
        sStr = sStr & xMatch.mHomeTeam & "     " & xMatch.mHomeTeamScore & "            " & xMatch.mAwayTeam & "     " & xMatch.mAwayTeamScore & vbCrLf
    Next
    
    GetScoreBoard = sStr
End Function
