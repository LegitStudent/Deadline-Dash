VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cStage5 
   Caption         =   "QuestKeeper"
   ClientHeight    =   8505.001
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11820
   OleObjectBlob   =   "cStage5.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cStage5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public currentDirection As String
Public stopMove As Boolean
Dim pogipoints As Integer
Dim startingTop As Integer
Dim startingLeft As Integer
Dim numberOfGuards As Integer
Dim numberOfBooks As Integer
Dim numberOfTrees As Integer
Dim numberOfHoles As Integer
Dim cheatsEnabled As Boolean
Dim timeLeft As Integer

'                     |=============================|
'                     |   ####         ###          |
'                     |  #            #   #         |
'                     |  #  harly's   #####         |
'                     |  #            #   #         |
'                     |   ####        #   #  ngels  |
'                     |=============================|
'
'                        D E A D L I N E  D A S H
'
'                                Francisco
'                                Lat
'                                Santiago
'                                Santos
'                                Tanjuatco

'==================================
' TIMING AND INITIAL FORM PROCEDURE
'==================================

'Description:

Private Sub UserForm_Activate() ' Main Game Loop, all code in this sub is executed every PauseTime seconds.
    Do While cStage5.Visible = True And stopMove = False
    PauseTime = 0.3 ' Set speed.
    Start = timer ' Set start time.
    
    Do While timer < Start + PauseTime
       DoEvents 'Yields to outside processes.
    Loop
    
    Select Case currentDirection
        Case "Down"
            Call cStage5.MoveDown
        Case "Up"
            Call cStage5.MoveUp
        Case "Left"
            Call cStage5.MoveLeft
        Case "Right"
            Call cStage5.MoveRight
    End Select
    
    
    Call GuardBlink
    If cheatsEnabled = False Then
        Call checkGameState
    End If
    
    'Time left
    '120     seconds
    
    threeTime = threeTime + 1
    
    If threeTime = 3 Then
        threeTime = 0
        timeLeft = timeLeft - 1
    End If
    
    If timeLeft = 1 Then
        cStage5.Caption = "ITM Lab - Deadline in " & timeLeft & " second"
    Else
        cStage5.Caption = "ITM Lab - Deadline in " & timeLeft & " seconds"
    End If
    
    Loop
End Sub

Private Sub UserForm_Initialize() ' Sets scroll bar to (0,0)
    cStage5.ScrollBars = fmScrollBarsNone
    cStage5.ScrollTop = 123
    cStage5.ScrollLeft = 320
    timeLeft = 120
    
' -----------------------------------------------------
' IMPORTANT: Remember to set your variables correctly.
' -----------------------------------------------------
    startingTop = 210
    startingLeft = 450
    numberOfHoles = 204
    numberOfGuards = 19
    numberOfBooks = 3
    numberOfTrees = 32
    
    Call RandomizeSir
End Sub

'===============
' USER CONTROLS
'===============

'Description: The last keypress is stored in a variable because change in direction is persistent.

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyUp Then
        currentDirection = "Up"
    ElseIf KeyCode = vbKeyDown Then
        currentDirection = "Down"
    ElseIf KeyCode = vbKeyLeft Then
        currentDirection = "Left"
    ElseIf KeyCode = vbKeyRight Then
        currentDirection = "Right"
    ElseIf KeyCode = vbKeyC Then
        cheatsEnabled = Not cheatsEnabled
    End If
    
End Sub

'============================
' PLAYER AND CAMERA MOVEMENT
'============================

'Description: Takes the currentDirection variable and moves the player in that direction whenever the procedure is called.
'             Also, if there is a tree in the way of the player or if the player is on the level's boundaries, it doesn't
'             let the player move.

Sub MoveUp()
    If charPlayer.Top = 0 Or checktree("Up") Then  ' top (player stops moving)
        'No code means nothing's happening
    Else
        charPlayer.Top = charPlayer.Top - 30
        cStage5.ScrollTop = charPlayer.Top - 87
    End If
End Sub

Sub MoveDown()
    If charPlayer.Top = 480 Or checktree("Down") Then    ' bottom (player stops moving)
    
    Else
        charPlayer.Top = charPlayer.Top + 30
        cStage5.ScrollTop = charPlayer.Top - 87
    End If
End Sub

Sub MoveRight()
    If charPlayer.Left = 690 Or checktree("Right") Then   ' rightmost side (player stops moving)
    
    Else
        charPlayer.Left = charPlayer.Left + 30
        cStage5.ScrollLeft = charPlayer.Left - 130
    End If
End Sub

Sub MoveLeft()
    If charPlayer.Left = 0 Or checktree("Left") Then     ' leftmost side (player stops moving)
    
    Else
        charPlayer.Left = charPlayer.Left - 30
        cStage5.ScrollLeft = charPlayer.Left - 130
    End If
End Sub

'======================================
' LOSE/WIN/POINTS - Collision Handling
'======================================

'Description: If a player collides with any object in the game, this procedure says what happens.
'             Collision is defined by getting the player in the same place/position as the other object.

Sub checkGameState()

'Hole collision
    For i = 1 To numberOfHoles
        If charPlayer.Left = Controls("Image" & i).Left And charPlayer.Top = Controls("Image" & i).Top Then
            Call endGame("loseHole")
        End If
    Next i
    
'Guard collision
    For i = 1 To numberOfGuards
        If charPlayer.Left = Controls("Guard" & i).Left And charPlayer.Top = Controls("Guard" & i).Top And Controls("Guard" & i).Visible Then
            Call endGame("loseGuard")
        End If
    Next i
    
'Book coliision, in ITM sir collision
    For i = 1 To numberOfBooks
        If charPlayer.Left = Controls("Book" & i).Left And charPlayer.Top = Controls("Book" & i).Top And Controls("Book" & i).Visible = True Then
            Call endGame("winSir")
        End If
    Next i
    
'Time Left
    If timeLeft = 0 Then
        Call endGame("loseTime")
    End If
End Sub

'=============================
' ASSORTED FUNCTIONS AND SUBS
'=============================

Function checktree(try_move As String) As Boolean

    'Description: This function looks ahead to see if the space the player is moving into is empty or if it has a tree.
    
    For i = 1 To numberOfTrees
        Select Case try_move
            Case Is = "Up"
                If charPlayer.Top - 30 = Controls("Tree" & i).Top And charPlayer.Left = Controls("Tree" & i).Left Then
                    checktree = True
                End If
            Case Is = "Down"
                If charPlayer.Top + 30 = Controls("Tree" & i).Top And charPlayer.Left = Controls("Tree" & i).Left Then
                    checktree = True
                End If
            Case Is = "Left"
                If charPlayer.Left - 30 = Controls("Tree" & i).Left And charPlayer.Top = Controls("Tree" & i).Top Then
                    checktree = True
                End If
            Case Is = "Right"
                If charPlayer.Left + 30 = Controls("Tree" & i).Left And charPlayer.Top = Controls("Tree" & i).Top Then
                    checktree = True
                End If
            End Select
    Next i
End Function

Sub GuardBlink()
    
    'Description: This procedure dictates the blinking of the guards. Blinking is based on probability.
    
    For i = 1 To numberOfGuards
        Randomize
        If Rnd() < 0.2 Then
            Controls("Guard" & i).Visible = True
            If Controls("Guard" & i).Top = charPlayer.Top And Controls("Guard" & i).Left = charPlayer.Left Then
                Call endGame("loseGuard") 'Fixes guard glitch
            End If
        Else
            Controls("Guard" & i).Visible = False
        End If
    Next i
End Sub

Sub endGame(condition As String)
    
    'Description: This procedure handles all winning and losing events.
    
    Select Case condition
        Case Is = "loseHole"
            MsgBox "You fell to your death!"
            cStage5.ScrollTop = 123
            cStage5.ScrollLeft = 320
            charPlayer.Left = startingLeft
            charPlayer.Top = startingTop
            currentDirection = ""
            Call RandomizeSir
            timeLeft = 120
            
        Case Is = "loseGuard"
            MsgBox "You got caught by Manong Guard!"
            charPlayer.Left = startingLeft
            charPlayer.Top = startingTop
            currentDirection = ""
            cStage5.ScrollTop = 123
            cStage5.ScrollLeft = 320
            Call RandomizeSir
            timeLeft = 120
            
        Case Is = "loseTime"
            MsgBox "You ran out of time!"
            charPlayer.Left = startingLeft
            charPlayer.Top = startingTop
            currentDirection = ""
            cStage5.ScrollTop = 123
            cStage5.ScrollLeft = 320
            Call RandomizeSir
            timeLeft = 120
            
        Case Is = "winSir"
            Unload cStage5
            TheEnd.Show
    End Select
End Sub

Sub RandomizeSir()
    Book1.Visible = False
    Book2.Visible = False
    Book3.Visible = False
    
    Randomize
    Select Case Rnd()
        Case Is < 0.33
            Book1.Visible = True
        Case Is < 0.66
            Book2.Visible = True
        Case Is <= 1
            Book3.Visible = True
    End Select
End Sub

