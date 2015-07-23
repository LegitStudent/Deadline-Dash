VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TheEnd 
   Caption         =   "The End"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11550
   OleObjectBlob   =   "TheEnd.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TheEnd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyDown <> vbKeyA Or KeyDown <> vbKeyB Then
        Unload Me
    End If
End Sub
