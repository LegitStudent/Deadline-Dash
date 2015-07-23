VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CutScene1 
   ClientHeight    =   7890
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12570
   OleObjectBlob   =   "CutScene1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CutScene1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Click()

End Sub

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode <> vbKeyA Or KeyCode <> vbKeyB Then
        Unload CutScene1
        instructions.Show
    End If
End Sub
