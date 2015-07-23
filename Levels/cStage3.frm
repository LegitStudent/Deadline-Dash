VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cStage3 
   Caption         =   "QuestKeeper"
   ClientHeight    =   8970.001
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12015
   OleObjectBlob   =   "cStage3.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cStage3"
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


'                   |=============================|
'                   |   ####         ###          |
'                   |  #            #   #         |
'                   |  #  harly's   #####         |
'                   |  #            #   #         |
'                   |   ####        #   #  ngels  |
'                   |=============================|
'
'                       D E A D L I N E  D A S H
'
'                               Francisco
'                               Lat
'                               Santiago
'                               Santos
'                               Tanjuatco

'==================================
' TIMING AND INITIAL FORM PROCEDURE
'==================================

'Description:

Private Sub UserForm_Activate() ' Main Game Loop, all code in this sub is executed every PauseTime seconds.
    Do While cStage3.Visible = True And stopMove = False
    PauseTime = 0.3 ' Set speed.
    Start = timer ' Set start time.
    
    Do While timer < Start + PauseTime
    DoEvents 'Yields to outside processes.
    Loop
    
    Select Case currentDirection
        Case "Down"
            Call cStage3.MoveDown
        Case "Up"
            Call cStage3.MoveUp
        Case "Left"
            Call cStage3.MoveLeft
        Case "Right"
            Call cStage3.MoveRight
    End Select
    
    
    Call GuardBlink
    If cheatsEnabled = False Then
        Call checkGameState
    End If
    
   cStage3.Caption = "Zen Garden - Books read: " & pogipoints & "/" & numberOfBooks
    
    Loop
End Sub

Private Sub UserForm_Initialize() ' Sets scroll bar to (0,0)
    cStage3.ScrollBars = fmScrollBarsNone
    cStage3.ScrollTop = 0
    cStage3.ScrollLeft = 0
    
' -----------------------------------------------------
' IMPORTANT: Remember to set your variables correctly.
' -----------------------------------------------------
    startingTop = 0
    startingLeft = 60
    numberOfHoles = 88
    numberOfGuards = 5
    numberOfBooks = 3
    numberOfTrees = 9
    
        
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
'           Also, if there is a tree in the way of the player or if the player is on the level's boundaries, it doesn't
'           let the player move.

Sub MoveUp()
    If charPlayer.Top = 0 Or checktree("Up") Then  ' top (player stops moving)
        'No code means nothing's happening
    Else
        charPlayer.Top = charPlayer.Top - 30
        If charPlayer.Top >= 100 Then
            cStage3.ScrollTop = charPlayer.Top - 100
        End If
    End If
End Sub

Sub MoveDown()
    If charPlayer.Top = 420 Or checktree("Down") Then   ' bottom (player stops moving)
    
    Else
        charPlayer.Top = charPlayer.Top + 30
        If charPlayer.Top >= 100 Then
            cStage3.ScrollTop = charPlayer.Top - 100
        End If
    End If
End Sub

Sub MoveRight()
    If charPlayer.Left = 570 Or checktree("Right") Then   ' rightmost side (player stops moving)
    
    Else
        charPlayer.Left = charPlayer.Left + 30
        If charPlayer.Left >= 100 Then
            cStage3.ScrollLeft = charPlayer.Left - 180
        End If
    End If
End Sub

Sub MoveLeft()
    If charPlayer.Left = 0 Or checktree("Left") Then    ' leftmost side (player stops moving)
    
    Else
        charPlayer.Left = charPlayer.Left - 30
        If charPlayer.Left >= 100 Then
            cStage3.ScrollLeft = charPlayer.Left - 180
        End If
    End If
End Sub

'======================================
' LOSE/WIN/POINTS - Collision Handling
'======================================

'Description: If a player collides with any object in the game, this procedure says what happens.
'           Collision is defined by getting the player in the same place/position as the other object.

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
    
'Book coliision
    For i = 1 To numberOfBooks
        If charPlayer.Left = Controls("Book" & i).Left And charPlayer.Top = Controls("Book" & i).Top And Controls("Book" & i).Visible = True Then
            Controls("Book" & i).Visible = False
            pogipoints = pogipoints + 1
        End If
     
        If pogipoints = numberOfBooks Then
            Call endGame("winBooks")
        End If
    Next i
    
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
            cStage3.ScrollLeft = 0
            cStage3.ScrollTop = 0
            charPlayer.Left = startingLeft
            charPlayer.Top = startingTop
            currentDirection = ""
         
            pogipoints = 0
            For j = 1 To numberOfBooks
                Controls("Book" & j).Visible = True
            Next j
         
        Case Is = "loseGuard"
            MsgBox "You got caught by Manong Guard!"
            charPlayer.Left = startingLeft
            charPlayer.Top = startingTop
            currentDirection = ""
            cStage3.ScrollLeft = 0
            cStage3.ScrollTop = 0
             
            ' reset all pogipoints
            pogipoints = 0
            For j = 1 To numberOfBooks
                Controls("Book" & j).Visible = True
            Next j
         
        Case Is = "winBooks"
            MsgBox "You got the books!"
            Unload cStage3
            sampleMenu.rbr_enabled = True
            sampleMenu.Show
         
    End Select
End Sub



