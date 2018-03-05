Attribute VB_Name = "Main"
Option Explicit

Enum t_City
    Belgorod
    Elk
    Chuguev
End Enum

Type t_SM
    radius As Double
    kfactor As Double
    matrix As Integer
    recommend As Boolean
    city As t_City
End Type

Type t_Sheet
    thickness As Double
    sm() As t_SM
End Type

Dim swApp As SldWorks.SldWorks
Dim swCurrentDoc As ModelDoc2  'maybe drawing
Dim swCurrentModel As ModelDoc2
Dim swCurrentDrawing As DrawingDoc
Dim swSheetMetalFeat As Feature
Dim swBaseFlangeFeat As Feature

Dim stdSheet(8) As t_Sheet
Dim currentThickness As Double
Dim currentRadius As Double
Dim currentKfactor As Double

Sub Main()
    Set swApp = Application.SldWorks
    Set swCurrentDoc = swApp.ActiveDoc
    
    If swCurrentDoc.GetType = swDocASSEMBLY Then
        Dim swDocSelMgr As SelectionMgr
        Set swDocSelMgr = swCurrentDoc.SelectionManager
        If swDocSelMgr.GetSelectedObjectCount2(-1) > 0 Then
            Set swCurrentModel = SelectedPart(swDocSelMgr.GetSelectedObjectsComponent3(1, -1).GetModelDoc2, "Выделите деталь")
        Else
            MsgBox "Выделите деталь", vbExclamation
            End
        End If
    ElseIf swCurrentDoc.GetType = swDocDRAWING Then
        Dim haveViews As Boolean
        Dim swCurrentSheet As sheet
        
        Set swCurrentDrawing = swCurrentDoc
        Set swCurrentSheet = swCurrentDrawing.GetCurrentSheet
        haveViews = False
        On Error Resume Next
        haveViews = UBound(swCurrentSheet.GetViews)
        If haveViews Then
            Dim swView As View
            Set swView = SelectView(swCurrentSheet)
            If swView.Type <> swDrawingStandardView Then
                Set swCurrentModel = SelectedPart(swView.ReferencedDocument, "Выберите вид с деталью")
            Else
                MsgBox "Пустой вид", vbExclamation
                End
            End If
        Else
            MsgBox "Пустой чертеж", vbExclamation
            End
        End If
    Else
        Set swCurrentModel = swCurrentDoc
    End If
    
    FindSheetMetal swCurrentModel
    If swSheetMetalFeat Is Nothing Or swBaseFlangeFeat Is Nothing Then
        MsgBox "Не найден листовой металл!", vbCritical
        Exit Sub
    End If
    'CheckSheetMetal  'maybe not necessary
    
    Dim swSheetMetal As Object
    Set swSheetMetal = swSheetMetalFeat.GetDefinition
    currentThickness = Round(swSheetMetal.thickness, 5)
    currentRadius = swSheetMetal.BendRadius
    currentKfactor = swSheetMetal.kfactor

    Init
    MainForm.Show
End Sub

Sub InitSheet(n_sheet As Integer, n_sm As Integer, radius As Double, kfactor As Double, _
               matrix As Integer, city As t_City, Optional recommend As Boolean = False)
    stdSheet(n_sheet).sm(n_sm).radius = radius / 1000
    stdSheet(n_sheet).sm(n_sm).kfactor = kfactor
    stdSheet(n_sheet).sm(n_sm).matrix = matrix
    stdSheet(n_sheet).sm(n_sm).city = city
    stdSheet(n_sheet).sm(n_sm).recommend = recommend
End Sub

Function Init()  'mask for button
    stdSheet(0).thickness = 0.001
    ReDim stdSheet(0).sm(6)
    InitSheet 0, 0, 2.09, 0.476, 8, Belgorod, True
    InitSheet 0, 1, 2.5, 0.422, 16, Elk, True
    InitSheet 0, 2, 3.38, 0.459, 22, Elk
    InitSheet 0, 3, 5.49, 0.501, 35, Elk
    InitSheet 0, 4, 7.92, 0.502, 50, Elk
    InitSheet 0, 5, 9.98, 0.499, 60, Elk
    InitSheet 0, 6, 12.43, 0.5, 80, Elk
    
    stdSheet(1).thickness = 0.0015
    ReDim stdSheet(1).sm(6)
    InitSheet 1, 0, 2.09, 0.338, 8, Belgorod, True
    InitSheet 1, 1, 2.5, 0.379, 16, Elk, True
    InitSheet 1, 2, 3.38, 0.412, 22, Elk
    InitSheet 1, 3, 5.49, 0.465, 35, Elk
    InitSheet 1, 4, 7.92, 0.501, 50, Elk
    InitSheet 1, 5, 9.98, 0.498, 60, Elk
    InitSheet 1, 6, 12.43, 0.499, 80, Elk
    
    stdSheet(2).thickness = 0.002
    ReDim stdSheet(2).sm(6)
    InitSheet 2, 0, 2.47, 0.369, 12, Belgorod, True
    InitSheet 2, 1, 2.5, 0.348, 16, Elk, True
    InitSheet 2, 2, 3.38, 0.382, 22, Elk
    InitSheet 2, 3, 5.49, 0.435, 35, Elk
    InitSheet 2, 4, 7.92, 0.474, 50, Elk
    InitSheet 2, 5, 9.98, 0.498, 60, Elk
    InitSheet 2, 6, 12.43, 0.501, 80, Elk

    stdSheet(3).thickness = 0.003
    ReDim stdSheet(3).sm(6)
    InitSheet 3, 0, 2.34, 0.319, 16, Belgorod, True
    InitSheet 3, 1, 2.5, 0.306, 16, Elk, True
    InitSheet 3, 2, 3.38, 0.338, 22, Elk
    InitSheet 3, 3, 5.49, 0.39, 35, Elk
    InitSheet 3, 4, 7.92, 0.431, 50, Elk
    InitSheet 3, 5, 9.98, 0.455, 60, Elk
    InitSheet 3, 6, 12.43, 0.481, 80, Elk
    
    stdSheet(4).thickness = 0.004
    ReDim stdSheet(4).sm(5)
    InitSheet 4, 0, 3.78, 0.298, 24, Belgorod, True
    InitSheet 4, 1, 3.38, 0.307, 22, Elk, True
    InitSheet 4, 2, 5.49, 0.359, 35, Elk
    InitSheet 4, 3, 7.92, 0.399, 50, Elk
    InitSheet 4, 4, 9.98, 0.424, 60, Elk
    InitSheet 4, 5, 12.43, 0.448, 80, Elk
    
    stdSheet(5).thickness = 0.005
    ReDim stdSheet(5).sm(4)
    InitSheet 5, 0, 6.2, 0.377, 32, Belgorod, True
    InitSheet 5, 1, 5.49, 0.336, 35, Elk, True
    InitSheet 5, 2, 7.92, 0.376, 50, Elk
    InitSheet 5, 3, 9.98, 0.4, 60, Elk
    InitSheet 5, 4, 12.43, 0.425, 80, Elk
    
    stdSheet(6).thickness = 0.006
    ReDim stdSheet(6).sm(3)
    InitSheet 6, 0, 7.39, 0.363, 40, Belgorod, True
    InitSheet 6, 1, 7.92, 0.355, 50, Elk, True
    InitSheet 6, 2, 9.98, 0.38, 60, Elk
    InitSheet 6, 3, 12.43, 0.404, 80, Elk
    
    stdSheet(7).thickness = 0.008
    ReDim stdSheet(7).sm(1)
    InitSheet 7, 0, 11.23, 0.368, 60, Belgorod, True
    InitSheet 7, 1, 12.43, 0.373, 80, Elk, True
    
    stdSheet(8).thickness = 0.01
    ReDim stdSheet(8).sm(0)
    InitSheet 8, 0, 14.89, 0.464, 80, Belgorod, True
    
    InitMainForm  'only after initialize stdSheet!
End Function

Function FindFeatureThisType(typeName As String, model As ModelDoc2) As Feature
    Dim feat As Feature
    
    Set feat = model.FirstFeature
    Do While Not feat Is Nothing
        If feat.GetTypeName2 = typeName Then
            Set FindFeatureThisType = feat
            Exit Do
        End If
        Set feat = feat.GetNextFeature
    Loop
End Function

'Magic function because of the bug in the API
'See more: https://forum.solidworks.com/thread/88666
Sub FindSheetMetal(model As ModelDoc2)
    Dim feat As Feature
    Dim swSheetMetalFolder As SheetMetalFolder
    
    Set swSheetMetalFolder = swCurrentModel.FeatureManager.GetSheetMetalFolder
    If swSheetMetalFolder Is Nothing Then  'for models created in SolidWorks 2012 and earlier
        Set swSheetMetalFeat = FindFeatureThisType("SheetMetal", model)
    Else
        Set swSheetMetalFeat = swSheetMetalFolder.GetFeature
    End If
    Set swBaseFlangeFeat = FindFeatureThisType("SMBaseFlange", model)
End Sub

Function CheckSheetMetal() As Boolean
    Dim swSheetMetal As SheetMetalFeatureData
    
    CheckSheetMetal = True
    Set swSheetMetal = swSheetMetalFeat.GetDefinition
    Select Case swSheetMetal.GetCustomBendAllowance.Type
        Case swBendAllowanceBendTable
            CheckSheetMetal = False
            MsgBox "Листовой металл управляется таблицей!", vbCritical
        Case Else
            CheckSheetMetal = False
            MsgBox "Неизвестный тип листового металла!", vbCritical
    End Select
End Function

Sub ChangeSheetMetal(s As Double, r As Double, k As Double)
    Dim swSheetMetal As SheetMetalFeatureData
    Dim swBaseFlange As IBaseFlangeFeatureData
    
    Set swBaseFlange = swBaseFlangeFeat.GetDefinition
    swBaseFlange.AccessSelections swCurrentModel, Nothing
    swBaseFlange.thickness = s
    swBaseFlangeFeat.ModifyDefinition swBaseFlange, swCurrentModel, Nothing
    
    Set swSheetMetal = swSheetMetalFeat.GetDefinition
    swSheetMetal.AccessSelections swCurrentModel, Nothing
    swSheetMetal.BendRadius = r
    swSheetMetal.kfactor = k
    swSheetMetalFeat.ModifyDefinition swSheetMetal, swCurrentModel, Nothing
End Sub

Sub Apply(index_sm As Integer, indexOfSheet As Integer)
    Dim sm As t_SM
    Dim aSheet As t_Sheet
    
    aSheet = stdSheet(indexOfSheet)
    sm = aSheet.sm(index_sm)
    
    ChangeSheetMetal aSheet.thickness, sm.radius, sm.kfactor
    
    If swCurrentDoc.GetType <> swDocPART Then
        FixRollBack
        swCurrentDoc.ForceRebuild3 True
    End If
End Sub

Function FixRollBack()  'mask for button
    Dim opt As swSaveAsOptions_e
    Dim err As swFileLoadError_e
    
    swApp.ActivateDoc3 swCurrentModel.GetPathName, False, opt, err
    swCurrentModel.FeatureManager.EditRollback swMoveRollbackBarToEnd, ""
    swApp.CloseDoc swCurrentModel.GetPathName
End Function

Function FindView(swSheet As sheet) As View
    Dim propView As String
    Dim firstView As View
    
    propView = swSheet.CustomPropertyView
    Set firstView = swCurrentDrawing.GetFirstView.GetNextView
    Set FindView = firstView
    Do While FindView.GetName2 <> propView
        Set FindView = FindView.GetNextView
        If FindView Is Nothing Then
            Set FindView = firstView
            Exit Do
        End If
    Loop
End Function

Function SelectView(swSheet As sheet) As View
    Dim selected As Object
    
    Set selected = swCurrentDrawing.SelectionManager.GetSelectedObject5(1)
    If selected Is Nothing Then
        Set SelectView = FindView(swSheet)
    ElseIf Not TypeOf selected Is View Then
        Set SelectView = FindView(swSheet)
    Else
        Set SelectView = selected
    End If
End Function

Function SelectedPart(swProbeModel As ModelDoc2, textWarning As String) As PartDoc
    If swProbeModel.GetType = swDocPART Then
        Set SelectedPart = swProbeModel
    Else
        MsgBox textWarning, vbExclamation
        End
    End If
End Function

Sub ChangeListRadiuses(indexOfSheet As Integer)
    Dim i As Integer
    Dim line As String
    Dim city As String
    Dim recommend As String
    Dim sep As String
    Dim eq As String
    Dim sm As t_SM
    
    For i = LBound(stdSheet(indexOfSheet).sm) To UBound(stdSheet(indexOfSheet).sm)
        sm = stdSheet(indexOfSheet).sm(i)
        
        If sm.city = Belgorod Then
            city = "Белгород"
        ElseIf sm.city = Elk Then
            city = "Элк     "
        ElseIf sm.city = Chuguev Then
            city = "Чугуев  "
        End If
        
        If sm.recommend Then
            recommend = "(реком.)"
        Else
            recommend = ""
        End If
        
        sep = "    "
        eq = " = "
        
        line = city + _
               sep + "R" + eq + Format(sm.radius * 1000, "00.00") + _
               sep + "K" + eq + Format(sm.kfactor, "0.000") + _
               sep + "V" + eq + Format(sm.matrix, "00") + _
               sep + recommend
        
        MainForm.listSm.AddItem line
        
        If currentRadius = sm.radius And currentKfactor = sm.kfactor Then
            MainForm.listSm.selected(i) = True
        End If
    Next
End Sub

Function InitMainForm()  'mask for button
    Dim i As Integer
    Dim indexOfSheet As Integer
    Dim isStandardThickness As Boolean
    
    For i = LBound(stdSheet) To UBound(stdSheet)
        MainForm.cmbThick.AddItem 1000 * stdSheet(i).thickness
    Next
    
    isStandardThickness = False
    For indexOfSheet = LBound(stdSheet) To UBound(stdSheet)
        If currentThickness = stdSheet(indexOfSheet).thickness Then  ' если толщина детали соответствует стандартной
            isStandardThickness = True
            Exit For
        End If
    Next
    
    If isStandardThickness Then
        MainForm.cmbThick.ListIndex = indexOfSheet
    Else
        MainForm.labThickness.Caption = "Толщина металла" + Str(currentThickness * 1000) + " мм"
    End If
End Function
