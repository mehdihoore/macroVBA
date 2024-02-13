Sub GroupRowsBasedOnMultipleLevels()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long, j As Long
    Dim currentLevel As Long
    Dim levelStarts() As Long
    Dim maxLevel As Integer
    
    Set ws = ActiveSheet ' Modify if working with a specific sheet
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    maxLevel = 10 ' Adjust based on the maximum expected WBS level
    
    ReDim levelStarts(maxLevel)
    
    Application.ScreenUpdating = False ' Turn off screen updating to speed up the script
    
    ' Initialize level starts
    For j = 1 To maxLevel
        levelStarts(j) = 0
    Next j
    
    For i = 2 To lastRow + 1 ' Loop to lastRow + 1 to ensure the last group is included
        If i <= lastRow Then
            currentLevel = Len(ws.Cells(i, 1).Value) - Len(Replace(ws.Cells(i, 1).Value, ".", "")) + 1
        Else
            currentLevel = 0 ' Force grouping at the end
        End If
        
        ' Update level starts and group previous levels if necessary
        For j = currentLevel + 1 To maxLevel
            If levelStarts(j) > 0 And i > levelStarts(j) Then
                ws.Rows(levelStarts(j) & ":" & i - 1).Group
                levelStarts(j) = 0
            End If
        Next j
        
        If currentLevel > 0 Then
            If levelStarts(currentLevel) = 0 Then
                levelStarts(currentLevel) = i
            End If
        End If
    Next i
    
    Application.ScreenUpdating = True ' Turn screen updating back on
End Sub

