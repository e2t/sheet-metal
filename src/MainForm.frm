VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainForm 
   Caption         =   "Sheet Metal"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8610.001
   OleObjectBlob   =   "MainForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub RunAndExit()
    Apply Me.listSm.ListIndex, Me.LstThickness.ListIndex
    ExitApp
End Sub

Private Sub butClose_Click()
    ExitApp
End Sub

Private Sub butRun_Click()
    RunAndExit
End Sub

Private Sub LstThickness_Change()
    Me.listSm.Clear
    Me.butRun.Enabled = False
    ChangeListRadiuses Me.LstThickness.ListIndex
End Sub

Private Sub listSm_Click()
    Me.butRun.Enabled = True
End Sub

Private Sub listSm_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    RunAndExit
End Sub

Private Sub LstThickness_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    RunAndExit
End Sub

Private Sub settingBut_Click()
    EditConfigFile
    ExitApp
End Sub

Private Sub UserForm_Initialize()
    Me.Caption = TitleWindow(Me.Caption)
End Sub
