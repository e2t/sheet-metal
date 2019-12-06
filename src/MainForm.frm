VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainForm 
   Caption         =   "Стандартные радиусы гибов"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6435
   OleObjectBlob   =   "MainForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub butClose_Click()
    End
End Sub

Private Sub butRun_Click()
    Apply listSm.ListIndex, cmbThick.ListIndex
    End
End Sub

Private Sub cmbThick_Change()
    listSm.Clear
    butRun.Enabled = False
    ChangeListRadiuses cmbThick.ListIndex
End Sub

Private Sub listSm_Click()
    butRun.Enabled = True
End Sub

Private Sub listSm_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Apply listSm.ListIndex, cmbThick.ListIndex
    End
End Sub

Private Sub settingBut_Click()
    EditConfigFile
    End
End Sub
