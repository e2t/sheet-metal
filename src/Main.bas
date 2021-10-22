Attribute VB_Name = "Main"
Option Explicit

Const ConfigFileName = "SheetMetal.conf"

Public swApp As SldWorks.SldWorks
Public gFSO As FileSystemObject

Dim gCurrentDoc As ModelDoc2  'maybe drawing
Dim gCurrentModel As ModelDoc2
Dim gSmMgr As TSheetMetalManager
Dim gStdSheets As Collection 'TSheet
Dim gConfigFullFileName As String

Sub Main()
  
  Dim HaveViews As Boolean
  Dim CurrentSheet As Sheet
  Dim DocSelMgr As SelectionMgr
  Dim AView As View
  Dim gCurrentDrawing As DrawingDoc

  Set swApp = Application.SldWorks
  Set gFSO = New FileSystemObject
  Set gCurrentDoc = swApp.ActiveDoc
  gConfigFullFileName = gFSO.BuildPath(swApp.GetCurrentMacroPathFolder, ConfigFileName)
  Set gStdSheets = New Collection
  
  If gCurrentDoc Is Nothing Then Exit Sub
  If gCurrentDoc.GetType = swDocASSEMBLY Then
    Set DocSelMgr = gCurrentDoc.SelectionManager
    If DocSelMgr.GetSelectedObjectCount2(-1) > 0 Then
      Set gCurrentModel = SelectedPart(DocSelMgr.GetSelectedObjectsComponent3(1, -1).GetModelDoc2, "Выделите деталь")
    Else
      MsgBox "Выделите деталь", vbExclamation
      End
    End If
  ElseIf gCurrentDoc.GetType = swDocDRAWING Then
    Set gCurrentDrawing = gCurrentDoc
    Set CurrentSheet = gCurrentDrawing.GetCurrentSheet
    HaveViews = False
    On Error Resume Next
    HaveViews = UBound(CurrentSheet.GetViews)
    If HaveViews Then
      Set AView = SelectView(CurrentSheet, gCurrentDrawing)
      If AView.Type <> swDrawingStandardView Then
        Set gCurrentModel = SelectedPart(AView.ReferencedDocument, "Выберите вид с деталью")
      Else
        MsgBox "Пустой вид", vbExclamation
        End
      End If
    Else
      MsgBox "Пустой чертеж", vbExclamation
      End
    End If
  Else
    Set gCurrentModel = gCurrentDoc
  End If
  
  Set gSmMgr = New TSheetMetalManager
  gSmMgr.Init gCurrentModel
  
  GetRowsFromFile gConfigFullFileName, gStdSheets
  InitMainForm
  MainForm.Show
    
End Sub

Sub Apply(IndexSm As Integer, IndexOfSheet As Integer)

  Dim Sm As TSm
  Dim ASheet As TSheet
  
  Set ASheet = gStdSheets(IndexOfSheet + 1)
  Set Sm = ASheet.Sm(IndexSm + 1)
  
  gCurrentModel.SetReadOnlyState False  'must be first!
  
  gSmMgr.ChangeSheetMetal ASheet.Thickness, Sm.Radius, Sm.KFactor
  
  If gCurrentDoc.GetType <> swDocPART Then
    FixRollBack gCurrentModel, gCurrentDoc
    gCurrentDoc.ForceRebuild3 True
  End If
    
End Sub

Sub ChangeListRadiuses(IndexOfSheet As Integer)

  Const Sep As String = "    "
  Const Eq As String = " = "

  Dim I As Variant
  Dim Sm As TSm
  Dim IndexSelectedRaidus As Integer
  Dim Line As String
  Dim SelectedSheet As TSheet
  
  Set SelectedSheet = gStdSheets.Item(IndexOfSheet + 1) 'collection index from 1
  For Each I In SelectedSheet.Sm
    Set Sm = I
    Line = "R" + Eq + Format(Sm.Radius * 1000, "00.00") + _
           Sep + "K" + Eq + Format(Sm.KFactor, "0.000") + _
           Sep + Sm.Note
    MainForm.listSm.AddItem Line
  Next
  
  If gSmMgr.CurrentThickness = SelectedSheet.Thickness Then
    IndexSelectedRaidus = SearchCurrentRadius(SelectedSheet.Sm, gSmMgr.CurrentRadius, gSmMgr.CurrentKFactor)
    If IndexSelectedRaidus < 0 And MainForm.Visible Then
      IndexSelectedRaidus = 0
    End If
  Else
    IndexSelectedRaidus = 0
  End If
  MainForm.listSm.ListIndex = IndexSelectedRaidus
    
End Sub

Function EditConfigFile() 'mask for button

  Shell "notepad " & gConfigFullFileName, vbNormalFocus

End Function

Function InitMainForm()  'mask for button

  Dim I As Integer
  Dim IndexOfSheet As Integer
  Dim IsStandardThickness As Boolean
  
  For I = 1 To gStdSheets.Count
    MainForm.LstThickness.AddItem 1000 * gStdSheets.Item(I).Thickness
  Next
  
  IsStandardThickness = False
  For IndexOfSheet = 1 To gStdSheets.Count
    If gSmMgr.CurrentThickness = gStdSheets.Item(IndexOfSheet).Thickness Then  ' если толщина детали соответствует стандартной
      IsStandardThickness = True
      Exit For
    End If
  Next
  
  If IsStandardThickness Then
    MainForm.LstThickness.ListIndex = IndexOfSheet - 1
  Else
    MainForm.labThickness.Caption = "Толщина металла" + Str(gSmMgr.CurrentThickness * 1000) + " мм"
  End If
  
  MainForm.LstThickness.SetFocus
  MainForm.LstThickness.Enabled = Not gSmMgr.BaseFlangeFeat Is Nothing
    
End Function

Function ExitApp() 'hide

  Unload MainForm
  End

End Function
