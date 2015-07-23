VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} instructions 
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13305
   OleObjectBlob   =   "instructions.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "instructions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode <> vbKeyA Or KeyCode <> vbKeyB Then
        Unload instructions
        sampleMenu.Show
    End If
End Sub
