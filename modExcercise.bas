Attribute VB_Name = "modExcercise"
Option Explicit

Public mCollScoreBoard As Collection
Public mScoreBoard As ClScoreBoard

Public Function GetScoreBoard() As String
    If mCollScoreBoard.Count = 0 Then Exit Function
    
    SortScoreBoard
    
    Dim lngCt As Long
    Dim xMatch As ClMatchInfo
    Dim sStr As String
    
    For lngCt = 1 To mCollScoreBoard.Count
        Set xMatch = mCollScoreBoard(lngCt)
        sStr = sStr & xMatch.mHomeTeam & "     " & xMatch.mHomeTeamScore & "            " & xMatch.mAwayTeam & "     " & xMatch.mAwayTeamScore & "    " & xMatch.mStartTime & vbCrLf
    Next
    
    GetScoreBoard = sStr
End Function


Private Sub SortScoreBoard()
     'Using Bubble sort
     'Sorting it twice jut to clear the oth criterias.
    Dim lngCt1 As Long
    Dim lngCt2 As Long
    Dim xMatch1 As ClMatchInfo
    Dim xMatch2 As ClMatchInfo
    
    'First Sort by total number of goals.
    For lngCt1 = 1 To mCollScoreBoard.Count - 1
        Set xMatch1 = mCollScoreBoard(lngCt1)
        
        For lngCt2 = lngCt1 + 1 To mCollScoreBoard.Count
            Set xMatch2 = mCollScoreBoard(lngCt2)
            If (xMatch1.mHomeTeamScore + xMatch1.mAwayTeamScore) < (xMatch2.mHomeTeamScore + xMatch2.mAwayTeamScore) Then
                MoveIndexInCollection lngCt2, lngCt1
            End If
        Next
    Next
    
    'Second Sort by for teams having same goals but sort by time.
    For lngCt1 = 1 To mCollScoreBoard.Count - 1
        Set xMatch1 = mCollScoreBoard(lngCt1)
        
        For lngCt2 = lngCt1 + 1 To mCollScoreBoard.Count
            Set xMatch2 = mCollScoreBoard(lngCt2)
            If (xMatch1.mHomeTeamScore + xMatch1.mAwayTeamScore) = (xMatch2.mHomeTeamScore + xMatch2.mAwayTeamScore) Then
                If xMatch1.mStartTime < xMatch2.mStartTime Then
                    MoveIndexInCollection lngCt2, lngCt1
                End If
            End If
        Next
    Next
End Sub

Private Sub MoveIndexInCollection(ByVal IndexBefore As Long, ByVal IndexAfter As Long)
    Dim xMatchToMove As ClMatchInfo: Set xMatchToMove = mCollScoreBoard(IndexBefore)
    mCollScoreBoard.Remove IndexBefore
    mCollScoreBoard.Add xMatchToMove, xMatchToMove.mHomeTeam & "||" & xMatchToMove.mAwayTeam, IndexAfter
End Sub
