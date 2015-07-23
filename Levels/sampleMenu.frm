VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} sampleMenu 
   Caption         =   "Map to ITM Lab"
   ClientHeight    =   8670.001
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11235
   OleObjectBlob   =   "sampleMenu.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "sampleMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rbr_enabled As Boolean
Public sec_enabled As Boolean
Public itm_enabled As Boolean
Public jsec_enabled As Boolean
Public zen_enabled As Boolean

Private Sub showjsec_Click()
    sampleMenu.Hide
    Load cStage1
    cStage1.Show
End Sub

Private Sub showsec_click()
    sampleMenu.Hide
    Load cStage2
    cStage2.Show
End Sub

Private Sub showZen_click()
    sampleMenu.Hide
    Load cStage3
    cStage3.Show
End Sub

Private Sub showredbrick_Click()
    sampleMenu.Hide
    Load cStage4
    cStage4.Show
End Sub

Private Sub UserForm_Activate()
    showredbrick.Enabled = rbr_enabled
    showsec.Enabled = sec_enabled
    showZen.Enabled = zen_enabled
    showITM.Enabled = itm_enabled
    
End Sub

Private Sub showITM_Click()
    answer = MsgBox("Not bluffing, this stage is really hard and you might want to have a glass of water first. Do you wish to continue?", vbYesNo)
    If answer = vbYes Then
        MsgBox ("You reach the ITM lab but you can't find Sir anywhere! You only have two minutes before the deadline. Find Sir and pass your project.")
        sampleMenu.Hide
        Load cStage5
        cStage5.Show
    End If
End Sub
