Attribute VB_Name = "Snake"
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Dim GameOver As Boolean
Dim Right As Boolean
Dim Left As Boolean
Dim Up As Boolean
Dim Down As Boolean
Dim DropSpawned As Boolean
Dim RemoveEnd As Boolean

Dim DropX As Integer
Dim DropY As Integer
Dim Height As Integer
Dim Width As Integer
Dim SnakeBody(1 To 100, 1 To 2) As Integer
Dim TempSnake(1 To 100, 1 To 2) As Integer

Public Sub Game()
Call CreateField
Call InitalizeSnake
Do While (Not GameOver)
    'DoEvents
    Call CheckKeys
    Call Move
    Call IsGameOver
    Call RandomDrop
    If (Not GameOver) Then
        Call DrawSnake
        Sleep 100
    End If
Loop
End Sub

Public Sub KIGame()
Call CreateField
Call InitalizeSnake
Do While (Not GameOver)
    DoEvents
    Call WhatWouldTheKIDo
    Call Move
    Call IsGameOver
    Call RandomDrop
    If (Not GameOver) Then
        Call DrawSnake
        Sleep 100
    End If
Loop
End Sub

Private Sub WhatWouldTheKIDo()
    If DropSpawned = True Then
        If SnakeBody(1, 1) < DropX And Left = False And Right = False Then
            Call RightButton
        ElseIf SnakeBody(1, 1) > DropX And Right = False And Left = False Then
            Call LeftButton
        ElseIf SnakeBody(1, 2) < DropY And Up = False And Down = False Then
            Call DownButton
        ElseIf SnakeBody(1, 2) > DropY And Down = False And Up = False Then
            Call UpButton
        End If
        Call CrashCheck
    End If
End Sub

Private Sub CrashCheck()
    If Right = True Then
        For i = LBound(SnakeBody) + 1 To UBound(SnakeBody) - 1
            If SnakeBody(i, 1) = 0 Then
                 Exit For
             End If
            If (SnakeBody(i, 1) = SnakeBody(1, 1) + 1 And SnakeBody(i, 2) = SnakeBody(1, 2)) Then
                Call UpButton
                Call CrashCheck
            End If
        Next i
    ElseIf Left = True Then
        For i = LBound(SnakeBody) + 1 To UBound(SnakeBody) - 1
            If SnakeBody(i, 1) = 0 Then
                 Exit For
             End If
            If (SnakeBody(i, 1) = SnakeBody(1, 1) - 1 And SnakeBody(i, 2) = SnakeBody(1, 2)) Then
                Call DownButton
                Call CrashCheck
            End If
        Next i
    ElseIf Up = True Then
        For i = LBound(SnakeBody) + 1 To UBound(SnakeBody) - 1
            If SnakeBody(i, 1) = 0 Then
                 Exit For
             End If
            If (SnakeBody(i, 1) = SnakeBody(1, 1) And SnakeBody(i, 2) = SnakeBody(1, 2) - 1) Then
                Call LeftButton
                Call CrashCheck
            End If
        Next i
    Else
        For i = LBound(SnakeBody) + 1 To UBound(SnakeBody) - 1
            If SnakeBody(i, 1) = 0 Then
                 Exit For
             End If
            If (SnakeBody(i, 1) = SnakeBody(1, 1) And SnakeBody(i, 2) = SnakeBody(1, 2) + 1) Then
                Call RightButton
                Call CrashCheck
            End If
        Next i
    End If
    
End Sub

Private Sub CreateField()
    LastScore = Worksheets("Settings").Cells(9, 6)
    HighScore = Worksheets("Settings").Cells(10, 6)
    Cells.Clear
    Height = Worksheets("Settings").Cells(9, 3)
    Width = Worksheets("Settings").Cells(10, 3)
    ActiveWindow.DisplayGridlines = False
    Worksheets("Settings").Cells(9, 6) = 0
    If LastScore > HighScore Then
        Worksheets("Settings").Cells(10, 6) = LastScore
    Else
        Worksheets("Settings").Cells(10, 6) = HighScore
    End If
    
    
    Columns("A:" & ColLetter(Width + 2)).ColumnWidth = 1.5
    Rows("1:" & Height + 2).RowHeight = 10.5
    
    With Range("A1:" & ColLetter(Width + 2) & "1").Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Range("A1:A" & Height + 2).Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Range("A" & Height + 2 & ":" & ColLetter(Width + 2) & Height + 2).Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Range(ColLetter(Width + 2) & "1:" & ColLetter(Width + 2) & Height + 2).Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub

Private Sub RandomDrop()
    If DropSpawned = False Then
        PermittedPosition = False
        Do While (Not PermittedPosition)
            PermittedPosition = True
            DropY = CInt((Height - 1) * Rnd() + 2)
            DropX = CInt((Width - 1) * Rnd() + 2)
            For i = LBound(SnakeBody) + 1 To UBound(SnakeBody)
                If SnakeBody(i, 1) = 0 Then
                     Exit For
                 End If
                If (SnakeBody(i, 1) = DropX And SnakeBody(i, 2) = DropY) Then
                    PermittedPosition = False
                    Exit For
                End If
            Next i
        Loop
        DropSpawned = True
        With Range(ColLetter(DropX) & DropY).Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorDark2
            .TintAndShade = -0.249977111117893
            .PatternTintAndShade = 0
        End With
    End If
End Sub

Private Sub CheckKeys()
If GetAsyncKeyState(vbKeyUp) Then
    Call UpButton
ElseIf GetAsyncKeyState(vbKeyRight) Then
    Call RightButton
ElseIf GetAsyncKeyState(vbKeyLeft) Then
    Call LeftButton
ElseIf GetAsyncKeyState(vbKeyDown) Then
    Call DownButton
End If
End Sub

Private Sub DrawSnake()
   For i = LBound(SnakeBody) To UBound(SnakeBody)
        If SnakeBody(i, 1) = 0 Then
            Exit For
        End If
        With Range(ColLetter(SnakeBody(i, 1)) & SnakeBody(i, 2)).Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 5296274
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
    Next i
End Sub
Private Sub Move()
    TempX = SnakeBody(1, 1)
    TempY = SnakeBody(1, 2)
    If Up = True Then
        TempSnake(1, 1) = TempX
        TempSnake(1, 2) = TempY - 1
    ElseIf Right = True Then
        TempSnake(1, 1) = TempX + 1
        TempSnake(1, 2) = TempY
    ElseIf Down = True Then
        TempSnake(1, 1) = TempX
        TempSnake(1, 2) = TempY + 1
    ElseIf Left = True Then
        TempSnake(1, 1) = TempX - 1
        TempSnake(1, 2) = TempY
    End If
    For i = LBound(SnakeBody) + 1 To UBound(SnakeBody)
        If SnakeBody(i - 1, 1) = 0 Then
            Exit For
        End If
        TempSnake(i, 1) = SnakeBody(i - 1, 1)
        TempSnake(i, 2) = SnakeBody(i - 1, 2)
    Next i
    Call CopyArray
End Sub
Private Sub IsGameOver()
    For i = LBound(SnakeBody) + 1 To UBound(SnakeBody)
        If SnakeBody(i, 1) = 0 Then
             Exit For
         End If
        If (SnakeBody(i, 1) = SnakeBody(1, 1) And SnakeBody(i, 2) = SnakeBody(1, 2)) Then
            Call GameIsOver
        End If
        If (SnakeBody(i, 1) = 1 Or SnakeBody(i, 1) = Width + 2 Or SnakeBody(i, 2) = 1 Or SnakeBody(i, 2) = Height + 2) Then
            Call GameIsOver
        End If
    Next i
    If SnakeBody(1, 1) = DropX And SnakeBody(1, 2) = DropY Then
        RemoveEnd = False
        DropSpawned = False
        Worksheets("Settings").Cells(9, 6) = Worksheets("Settings").Cells(9, 6) + 1
    End If
End Sub
Private Sub CopyArray()
For i = LBound(TempSnake) To UBound(TempSnake)
        If TempSnake(i, 1) = 0 Then
            If i = 1 Then
                Call GameIsOver
            Exit Sub
            Else
                If RemoveEnd = True Then
                    With Cells(SnakeBody(i - 1, 2), SnakeBody(i - 1, 1)).Interior
                        .Pattern = xlNone
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    SnakeBody(i - 1, 1) = 0
                    SnakeBody(i - 1, 2) = 0
                    TempSnake(i - 1, 1) = 0
                    TempSnake(i - 1, 2) = 0
                End If
                RemoveEnd = True
                Exit For
            End If
        End If
        SnakeBody(i, 1) = TempSnake(i, 1)
        SnakeBody(i, 2) = TempSnake(i, 2)
    Next i
End Sub
Private Sub GameIsOver()
    GameOver = True
    MsgBox "Game Over! Score: " & Worksheets("Settings").Cells(9, 6) & "."
End Sub

Private Sub InitalizeSnake()
GameOver = False
Left = False
Call RightButton
DropSpawned = False
RemoveEnd = True
DropX = 0
DropY = 0

For i = LBound(SnakeBody) + 1 To UBound(SnakeBody)
    SnakeBody(i, 1) = 0
    SnakeBody(i, 2) = 0
    TempSnake(i, 1) = 0
    TempSnake(i, 2) = 0
Next i

SnakeBody(1, 1) = 20
SnakeBody(1, 2) = 14
SnakeBody(2, 1) = 19
SnakeBody(2, 2) = 14
SnakeBody(3, 1) = 18
SnakeBody(3, 2) = 14
SnakeBody(4, 1) = 17
SnakeBody(4, 2) = 14
SnakeBody(5, 1) = 16
SnakeBody(5, 2) = 14
SnakeBody(6, 1) = 15
SnakeBody(6, 2) = 14
Call DrawSnake
End Sub

Private Sub UpButton()
    If Down <> True Then
        Up = True
        Down = False
        Right = False
        Left = False
    End If
End Sub
Private Sub RightButton()
    If Left <> True Then
        Up = False
        Down = False
        Right = True
        Left = False
    End If
End Sub
Private Sub DownButton()
    If Up <> True Then
        Up = False
        Down = True
        Right = False
        Left = False
    End If
End Sub
Private Sub LeftButton()
    If Right <> True Then
        Up = False
        Down = False
        Right = False
        Left = True
    End If
End Sub
Function ColLetter(iCol As Integer) As String
   Dim iAlpha As Integer
   Dim iRemainder As Integer
   iAlpha = Int(iCol / 27)
   iRemainder = iCol - (iAlpha * 26)
   If iAlpha > 0 Then
      ColLetter = Chr(iAlpha + 64)
   End If
   If iRemainder > 0 Then
      ColLetter = ColLetter & Chr(iRemainder + 64)
   End If
End Function

