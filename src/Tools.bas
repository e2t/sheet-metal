Attribute VB_Name = "Tools"
Option Explicit

Function TitleWindow(Optional text As String = "") As String
    If text = "" Then
        TitleWindow = MacroName + " " + MacroVersion
    Else
        TitleWindow = text + " — " + MacroName + " " + MacroVersion
    End If
End Function

Sub CreateDefaultConfigFile(ObjStream As Stream, ConfigFullFileName As String)
    'TODO: check if cannot to create file
    ObjStream.WriteText _
        "[1]" & vbNewLine & _
        "2.09  0.476 Белгород V=8, Полка 7" & vbNewLine & _
        "2.5   0.422 Польша V=16" & vbNewLine & _
        "3.38  0.459 Польша V=22" & vbNewLine & _
        "5.49  0.501 Польша V=35" & vbNewLine & _
        "7.92  0.502 Польша V=50" & vbNewLine & _
        "9.98  0.499 Польша V=60" & vbNewLine & _
        "12.43 0.5   Польша V=80" & vbNewLine & _
        vbNewLine
    ObjStream.WriteText _
        "[1.5]" & vbNewLine & _
        "2.09  0.338 Белгород V=8, Полка 7" & vbNewLine & _
        "2.5   0.379 Польша V=16" & vbNewLine & _
        "3.38  0.412 Польша V=22" & vbNewLine & _
        "5.49  0.465 Польша V=35" & vbNewLine & _
        "7.92  0.501 Польша V=50" & vbNewLine & _
        "9.98  0.498 Польша V=60" & vbNewLine & _
        "12.43 0.499 Польша V=80" & vbNewLine & _
        vbNewLine
    ObjStream.WriteText _
        "[2]" & vbNewLine & _
        "2.47  0.369 Белгород V=12, Полка 10" & vbNewLine & _
        "2.5   0.348 Польша V=16" & vbNewLine & _
        "3.38  0.382 Польша V=22" & vbNewLine & _
        "5.49  0.435 Польша V=35" & vbNewLine & _
        "7.92  0.474 Польша V=50" & vbNewLine & _
        "9.98  0.498 Польша V=60" & vbNewLine & _
        "12.43 0.501 Польша V=80" & vbNewLine & _
        vbNewLine
    ObjStream.WriteText _
        "[3]" & vbNewLine & _
        "2.34  0.319 Белгород V=16, Полка 13" & vbNewLine & _
        "2.5   0.306 Польша V=16" & vbNewLine & _
        "3.38  0.338 Польша V=22" & vbNewLine & _
        "5.49  0.39  Польша V=35" & vbNewLine & _
        "7.92  0.431 Польша V=50" & vbNewLine & _
        "9.98  0.455 Польша V=60" & vbNewLine & _
        "12.43 0.481 Польша V=80" & vbNewLine & _
        vbNewLine
    ObjStream.WriteText _
        "[4]" & vbNewLine & _
        "3.78  0.298 Белгород V=24, Полка 19" & vbNewLine & _
        "3.38  0.307 Польша V=22" & vbNewLine & _
        "5.49  0.359 Польша V=35" & vbNewLine & _
        "7.92  0.399 Польша V=50" & vbNewLine & _
        "9.98  0.424 Польша V=60" & vbNewLine & _
        "12.43 0.448 Польша V=80" & vbNewLine & _
        vbNewLine
    ObjStream.WriteText _
        "[5]" & vbNewLine & _
        "6.2   0.377 Белгород V=32, Полка 24" & vbNewLine & _
        "5.49  0.336 Польша V=35" & vbNewLine & _
        "7.92  0.376 Польша V=50" & vbNewLine & _
        "9.98  0.4   Польша V=60" & vbNewLine & _
        "12.43 0.425 Польша V=80" & vbNewLine & _
        vbNewLine
    ObjStream.WriteText _
        "[6]" & vbNewLine & _
        "7.39  0.363 Белгород V=40, Полка 29" & vbNewLine & _
        "7.92  0.355 Польша V=50" & vbNewLine & _
        "9.98  0.38  Польша V=60" & vbNewLine & _
        "12.43 0.404 Польша V=80" & vbNewLine & _
        vbNewLine
    ObjStream.WriteText _
        "[8]" & vbNewLine & _
        "11.23 0.368 Белгород V=60, Полка 42" & vbNewLine & _
        "12.43 0.373 Польша V=80" & vbNewLine & _
        vbNewLine
    ObjStream.WriteText _
        "[10]" & vbNewLine & _
        "14.89 0.464 Белгород V=80, Полка 56" & vbNewLine & _
        vbNewLine
    ObjStream.SaveToFile ConfigFullFileName
End Sub

Function FindFeatureThisType(TypeName As String, Model As ModelDoc2) As Feature
    Dim Feat As Feature
    
    Set Feat = Model.FirstFeature
    Do While Not Feat Is Nothing
        If Feat.GetTypeName2 = TypeName Then
            Set FindFeatureThisType = Feat
            Exit Do
        End If
        Set Feat = Feat.GetNextFeature
    Loop
End Function

Function FindView(ASheet As Sheet, CurrentDrawing As DrawingDoc) As View
    Dim PropView As String
    Dim FirstView As View
    
    PropView = ASheet.CustomPropertyView
    Set FirstView = CurrentDrawing.GetFirstView.GetNextView
    Set FindView = FirstView
    If Not FirstView Is Nothing Then
        Do While FindView.GetName2 <> PropView
            Set FindView = FindView.GetNextView
            If FindView Is Nothing Then
                Set FindView = FirstView
                Exit Do
            End If
        Loop
    End If
End Function

Function SelectView(ASheet As Sheet, CurrentDrawing As DrawingDoc) As View
    Dim Selected As Object
    
    Set Selected = CurrentDrawing.SelectionManager.GetSelectedObject5(1)
    If Selected Is Nothing Then
        Set SelectView = FindView(ASheet, CurrentDrawing)
    ElseIf Not TypeOf Selected Is View Then
        Set SelectView = FindView(ASheet, CurrentDrawing)
    Else
        Set SelectView = Selected
    End If
End Function

Function SelectedPart(swProbeModel As ModelDoc2, TextWarning As String) As PartDoc
    If swProbeModel.GetType = swDocPART Then
        Set SelectedPart = swProbeModel
    Else
        MsgBox TextWarning, vbExclamation
        End
    End If
End Function

Sub ReadRowsFromFile(ObjStream As Stream, ByRef StdSheets As Collection)
    Const RowIsSection As Integer = 1
    Const RowIsItem As Integer = 2
    
    Dim ASheet As TSheet
    Dim StrData As String
    Dim SectionMatch As MatchCollection
    Dim LineMatch As MatchCollection
    Dim Sm As TSm
    Dim SectionRegex As RegExp
    Dim LineRegex As RegExp
    
    Set SectionRegex = New RegExp
    SectionRegex.Pattern = "\[\s*([0-9.]+)(.*)\]"
    SectionRegex.IgnoreCase = True
    
    Set LineRegex = New RegExp
    LineRegex.Pattern = "([0-9.]+)\s+([0-9.]+)(.*)"
    LineRegex.IgnoreCase = True
    
    Do Until ObjStream.EOS
        StrData = Trim(ObjStream.ReadText(adReadLine))
        
        If Len(StrData) < 1 Then
            GoTo NextDo
        End If
      
        On Error GoTo NextDo
        If SectionRegex.Test(StrData) Then
            Set SectionMatch = SectionRegex.Execute(StrData)
            Set ASheet = New TSheet
            ASheet.Thickness = Val(SectionMatch.Item(0).SubMatches.Item(0)) / 1000#  ' mm => m
            Set ASheet.Sm = New Collection
            StdSheets.Add ASheet
        ElseIf LineRegex.Test(StrData) Then
            Set LineMatch = LineRegex.Execute(StrData)
            Set Sm = New TSm
            Sm.Radius = Val(LineMatch.Item(0).SubMatches.Item(0)) / 1000#  ' mm => m
            Sm.KFactor = Val(LineMatch.Item(0).SubMatches.Item(1))
            Sm.Note = Trim(LineMatch.Item(0).SubMatches.Item(2))
            StdSheets.Item(StdSheets.Count).Sm.Add Sm
        End If
        GoTo NextDo
NextDo:
    Loop
End Sub

Function GetRowsFromFile(ConfigFullFileName As String, ByRef StdSheets As Collection) As Boolean
    Dim ObjStream As Stream
        
    Set ObjStream = New Stream
    ObjStream.Charset = "utf-8"
    ObjStream.Open
    GetRowsFromFile = False
    
    On Error GoTo CreateConfig
    ObjStream.LoadFromFile ConfigFullFileName
    GoTo SuccessRead

ReadConfigAgain:
    On Error GoTo ExitFunction
    ObjStream.LoadFromFile ConfigFullFileName
    GoTo SuccessRead
   
SuccessRead:
    ReadRowsFromFile ObjStream, StdSheets
    GetRowsFromFile = True
  
ExitFunction:
    ObjStream.Close
    Set ObjStream = Nothing
    Exit Function
    
CreateConfig:
    CreateDefaultConfigFile ObjStream, ConfigFullFileName
    GoTo ReadConfigAgain
End Function

'Magic function because of the bug in the API
'See more: https://forum.solidworks.com/thread/88666
Sub FindSheetMetal( _
    Model As ModelDoc2, ByRef SheetMetalFeat As Feature, ByRef BaseFlangeFeat As Feature)
    
    Dim Feat As Feature
    Dim ASheetMetalFolder As SheetMetalFolder
    
    Set ASheetMetalFolder = Model.FeatureManager.GetSheetMetalFolder
    If ASheetMetalFolder Is Nothing Then  'for models created in SolidWorks 2012 and earlier
        Set SheetMetalFeat = FindFeatureThisType("SheetMetal", Model)
    Else
        Set SheetMetalFeat = ASheetMetalFolder.GetFeature
    End If
    Set BaseFlangeFeat = FindFeatureThisType("SMBaseFlange", Model)
End Sub

Function IsOpened(Model As ModelDoc2) As Boolean
    Dim I As Variant
    Dim AModelWindow As ModelWindow
    
    IsOpened = False
    For Each I In swApp.Frame.ModelWindows
        Set AModelWindow = I
        If AModelWindow.ModelDoc Is Model Then
            IsOpened = True
            Exit For
        End If
    Next
End Function

Function FixRollBack(CurrentModel As ModelDoc2, CurrentDoc As ModelDoc2)  'mask for button
    Dim Opt As swSaveAsOptions_e
    Dim Err As swFileLoadError_e
    Dim I As Variant
    Dim WasOpened As Boolean
    
    WasOpened = IsOpened(CurrentModel)
    swApp.ActivateDoc3 CurrentModel.GetPathName, False, Opt, Err
    CurrentModel.FeatureManager.EditRollback swMoveRollbackBarToEnd, ""
    If Not CurrentModel Is CurrentDoc Then
        If WasOpened Then
            swApp.ActivateDoc3 CurrentDoc.GetPathName, False, Opt, Err
        Else
            swApp.CloseDoc CurrentModel.GetPathName
        End If
    End If
    CurrentModel.SetSaveFlag
End Function

Function SearchCurrentRadius(Sm As Collection, R As Double, K As Double) As Integer
    Dim I As Integer
    
    SearchCurrentRadius = -1
    For I = 1 To Sm.Count
        If (R = Sm.Item(I).Radius) And (K = Sm.Item(I).KFactor) Then
            SearchCurrentRadius = I - 1
        End If
    Next
End Function
