Attribute VB_Name = "Main"
Option Explicit

Private Const configFileName As String = "SheetMetal.conf"

Dim swApp As SldWorks.SldWorks
Dim swCurrentDoc As ModelDoc2  'maybe drawing
Dim swCurrentModel As ModelDoc2
Dim swCurrentDrawing As DrawingDoc
Dim swSheetMetalFeat As Feature
Dim swBaseFlangeFeat As Feature

Dim stdSheet As Collection 'sheet_t
Dim currentThickness As Double
Dim currentRadius As Double
Dim currentKfactor As Double

Dim configFullFileName As String
Dim sectionRegex As RegExp
Dim lineRegex As RegExp
     
Sub Main()
    Set swApp = Application.SldWorks
    Set swCurrentDoc = swApp.ActiveDoc
    configFullFileName = swApp.GetCurrentMacroPathFolder + "\" + configFileName
    Set stdSheet = New Collection
     
    Set sectionRegex = New RegExp
    sectionRegex.Pattern = "\[\s*([0-9.]+)(.*)\]"
    sectionRegex.IgnoreCase = True
    
    Set lineRegex = New RegExp
    lineRegex.Pattern = "([0-9.]+)\s+([0-9.]+)(.*)"
    lineRegex.IgnoreCase = True
    
    If swCurrentDoc Is Nothing Then Exit Sub
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
    If swSheetMetalFeat Is Nothing Then
        MsgBox "Не найден листовой металл!", vbCritical
        Exit Sub
    End If
    
    Dim swSheetMetal As Object
    Set swSheetMetal = swSheetMetalFeat.GetDefinition
    currentThickness = Round(swSheetMetal.thickness, 5)
    currentRadius = swSheetMetal.BendRadius
    currentKfactor = swSheetMetal.kfactor
    
    GetRowsFromFile
    InitMainForm
    MainForm.Show
End Sub

Function GetRowsFromFile() As Boolean
    Dim objStream As Stream
        
    Set objStream = New Stream
    objStream.Charset = "utf-8"
    objStream.Open
    GetRowsFromFile = False
    
    On Error GoTo CreateConfig
    objStream.LoadFromFile configFullFileName
    GoTo SuccessRead

ReadConfigAgain:
    On Error GoTo ExitFunction
    objStream.LoadFromFile configFullFileName
    GoTo SuccessRead
   
SuccessRead:
    ReadRowsFromFile objStream
    GetRowsFromFile = True
ExitFunction:
    objStream.Close
    Set objStream = Nothing
    Exit Function
    
CreateConfig:
    CreateDefaultConfigFile objStream
    GoTo ReadConfigAgain
End Function

Sub ReadRowsFromFile(objStream As Stream)
    Const RowIsSection As Integer = 1
    Const RowIsItem As Integer = 2
    
    Dim asheet As sheet_t
    Dim strData
    Dim sectionMatch As MatchCollection
    Dim lineMatch As MatchCollection
    Dim sm As sm_t
    
    Do Until objStream.EOS
        strData = Trim(objStream.ReadText(adReadLine))
        
        If Len(strData) < 1 Then
            GoTo NextDo
        End If
        
        On Error GoTo NextDo
        If sectionRegex.Test(strData) Then
            Set sectionMatch = sectionRegex.Execute(strData)
            Set asheet = New sheet_t
            asheet.thickness = Val(sectionMatch.Item(0).SubMatches.Item(0)) / 1000#  ' mm => m
            Set asheet.sm = New Collection
            stdSheet.Add asheet
        ElseIf lineRegex.Test(strData) Then
            Set lineMatch = lineRegex.Execute(strData)
            Set sm = New sm_t
            sm.radius = Val(lineMatch.Item(0).SubMatches.Item(0)) / 1000#  ' mm => m
            sm.kfactor = Val(lineMatch.Item(0).SubMatches.Item(1))
            sm.note = Trim(lineMatch.Item(0).SubMatches.Item(2))
            'MsgBox Str(sm.radius) & "_" & Str(sm.kfactor) & "_" & sm.note
            stdSheet.Item(stdSheet.Count).sm.Add sm
        End If
        GoTo NextDo
NextDo:
    Loop
End Sub

Sub CreateDefaultConfigFile(objStream As Stream)
    'TODO: check if cannot to create file
    objStream.WriteText _
        "[1]" & vbNewLine & _
        "2.09  0.476 Белгород V=8" & vbNewLine & _
        "2.5   0.422 Польша V=16" & vbNewLine & _
        "3.38  0.459 Польша V=22" & vbNewLine & _
        "5.49  0.501 Польша V=35" & vbNewLine & _
        "7.92  0.502 Польша V=50" & vbNewLine & _
        "9.98  0.499 Польша V=60" & vbNewLine & _
        "12.43 0.5   Польша V=80" & vbNewLine & _
        vbNewLine
    objStream.WriteText _
        "[1.5]" & vbNewLine & _
        "2.09  0.338 Белгород V=8" & vbNewLine & _
        "2.5   0.379 Польша V=16" & vbNewLine & _
        "3.38  0.412 Польша V=22" & vbNewLine & _
        "5.49  0.465 Польша V=35" & vbNewLine & _
        "7.92  0.501 Польша V=50" & vbNewLine & _
        "9.98  0.498 Польша V=60" & vbNewLine & _
        "12.43 0.499 Польша V=80" & vbNewLine & _
        vbNewLine
    objStream.WriteText _
        "[2]" & vbNewLine & _
        "2.47  0.369 Белгород V=12" & vbNewLine & _
        "2.5   0.348 Польша V=16" & vbNewLine & _
        "3.38  0.382 Польша V=22" & vbNewLine & _
        "5.49  0.435 Польша V=35" & vbNewLine & _
        "7.92  0.474 Польша V=50" & vbNewLine & _
        "9.98  0.498 Польша V=60" & vbNewLine & _
        "12.43 0.501 Польша V=80" & vbNewLine & _
        vbNewLine
    objStream.WriteText _
        "[3]" & vbNewLine & _
        "2.34  0.319 Белгород V=16" & vbNewLine & _
        "2.5   0.306 Польша V=16" & vbNewLine & _
        "3.38  0.338 Польша V=22" & vbNewLine & _
        "5.49  0.39  Польша V=35" & vbNewLine & _
        "7.92  0.431 Польша V=50" & vbNewLine & _
        "9.98  0.455 Польша V=60" & vbNewLine & _
        "12.43 0.481 Польша V=80" & vbNewLine & _
        vbNewLine
    objStream.WriteText _
        "[4]" & vbNewLine & _
        "3.78  0.298 Белгород V=24" & vbNewLine & _
        "3.38  0.307 Польша V=22" & vbNewLine & _
        "5.49  0.359 Польша V=35" & vbNewLine & _
        "7.92  0.399 Польша V=50" & vbNewLine & _
        "9.98  0.424 Польша V=60" & vbNewLine & _
        "12.43 0.448 Польша V=80" & vbNewLine & _
        vbNewLine
    objStream.WriteText _
        "[5]" & vbNewLine & _
        "6.2   0.377 Белгород V=32" & vbNewLine & _
        "5.49  0.336 Польша V=35" & vbNewLine & _
        "7.92  0.376 Польша V=50" & vbNewLine & _
        "9.98  0.4   Польша V=60" & vbNewLine & _
        "12.43 0.425 Польша V=80" & vbNewLine & _
        vbNewLine
    objStream.WriteText _
        "[6]" & vbNewLine & _
        "7.39  0.363 Белгород V=40" & vbNewLine & _
        "7.92  0.355 Польша V=50" & vbNewLine & _
        "9.98  0.38  Польша V=60" & vbNewLine & _
        "12.43 0.404 Польша V=80" & vbNewLine & _
        vbNewLine
    objStream.WriteText _
        "[8]" & vbNewLine & _
        "11.23 0.368 Белгород V=60" & vbNewLine & _
        "12.43 0.373 Польша V=80" & vbNewLine & _
        vbNewLine
    objStream.WriteText _
        "[10]" & vbNewLine & _
        "14.89 0.464 Белгород V=80" & vbNewLine & _
        vbNewLine
    objStream.SaveToFile configFullFileName
End Sub

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
    
    If Not swBaseFlangeFeat Is Nothing Then
        Set swBaseFlange = swBaseFlangeFeat.GetDefinition
        swBaseFlange.AccessSelections swCurrentModel, Nothing
        swBaseFlange.thickness = s
        swBaseFlangeFeat.ModifyDefinition swBaseFlange, swCurrentModel, Nothing
    End If
    
    Set swSheetMetal = swSheetMetalFeat.GetDefinition
    swSheetMetal.AccessSelections swCurrentModel, Nothing
    swSheetMetal.BendRadius = r
    swSheetMetal.kfactor = k
    swSheetMetalFeat.ModifyDefinition swSheetMetal, swCurrentModel, Nothing
End Sub

Sub Apply(index_sm As Integer, indexOfSheet As Integer)
    Dim sm As sm_t
    Dim asheet As sheet_t
    
    Set asheet = stdSheet(indexOfSheet + 1)
    Set sm = asheet.sm(index_sm + 1)
    
    ChangeSheetMetal asheet.thickness, sm.radius, sm.kfactor
    
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
    Dim i As Variant
    
    Dim selectedSheet As sheet_t
    Set selectedSheet = stdSheet.Item(indexOfSheet + 1) 'collection index from 1
    For Each i In selectedSheet.sm
        Dim sm As sm_t
        Set sm = i
        
        Const sep As String = "    "
        Const eq As String = " = "
    
        Dim line As String
        line = "R" + eq + Format(sm.radius * 1000, "00.00") + _
               sep + "K" + eq + Format(sm.kfactor, "0.000") + _
               sep + sm.note
        
        MainForm.listSm.AddItem line
    Next
    
    If currentThickness = selectedSheet.thickness Then
        Dim indexSelectedRaidus As Integer
        indexSelectedRaidus = SearchCurrentRadius(selectedSheet.sm)
        If indexSelectedRaidus < 0 And MainForm.Visible Then
            indexSelectedRaidus = 0
        End If
    Else
        indexSelectedRaidus = 0
    End If
    MainForm.listSm.ListIndex = indexSelectedRaidus
End Sub

Function SearchCurrentRadius(sm As Collection) As Integer
    Dim i As Integer
    
    SearchCurrentRadius = -1
    For i = 1 To sm.Count
        If currentRadius = sm.Item(i).radius And currentKfactor = sm.Item(i).kfactor Then
            SearchCurrentRadius = i - 1
        End If
    Next
End Function

Function EditConfigFile() 'mask for button
    Shell "notepad " & configFullFileName, vbNormalFocus
End Function

Function InitMainForm()  'mask for button
    Dim i As Integer
    Dim indexOfSheet As Integer
    Dim isStandardThickness As Boolean
    
    For i = 1 To stdSheet.Count
        MainForm.cmbThick.AddItem 1000 * stdSheet.Item(i).thickness
    Next
    
    isStandardThickness = False
    For indexOfSheet = 1 To stdSheet.Count
        If currentThickness = stdSheet.Item(indexOfSheet).thickness Then  ' если толщина детали соответствует стандартной
            isStandardThickness = True
            Exit For
        End If
    Next
    
    If isStandardThickness Then
        MainForm.cmbThick.ListIndex = indexOfSheet - 1
    Else
        MainForm.labThickness.Caption = "Толщина металла" + Str(currentThickness * 1000) + " мм"
    End If
    
    MainForm.cmbThick.Enabled = Not swBaseFlangeFeat Is Nothing
End Function
