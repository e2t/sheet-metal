VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TSheetMetalManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim CurrentModel As ModelDoc2
Public SheetMetalFeat As Feature
Public BaseFlangeFeat As Feature
Public CurrentThickness As Double
Public CurrentRadius As Double
Public CurrentKFactor As Double

Sub Init(ACurrentModel As ModelDoc2)

  Dim SheetMetal As Object

  Set CurrentModel = ACurrentModel
  FindSheetMetal CurrentModel, SheetMetalFeat, BaseFlangeFeat
  If SheetMetalFeat Is Nothing Then
    MsgBox "Не найден листовой металл!", vbCritical
    ExitApp
  End If
  Set SheetMetal = SheetMetalFeat.GetDefinition
  CurrentThickness = Round(SheetMetal.Thickness, 5)
  CurrentRadius = SheetMetal.BendRadius
  CurrentKFactor = SheetMetal.KFactor

End Sub

Sub ChangeSheetMetal(S As Double, R As Double, K As Double)

  Dim SheetMetal As SheetMetalFeatureData
  Dim BaseFlange As IBaseFlangeFeatureData

  If Not BaseFlangeFeat Is Nothing Then
    Set BaseFlange = BaseFlangeFeat.GetDefinition
    BaseFlange.AccessSelections CurrentModel, Nothing
    BaseFlange.Thickness = S
    BaseFlangeFeat.ModifyDefinition BaseFlange, CurrentModel, Nothing
  End If
  
  Set SheetMetal = SheetMetalFeat.GetDefinition
  SheetMetal.AccessSelections CurrentModel, Nothing
  SheetMetal.BendRadius = R
  SheetMetal.KFactor = K
  SheetMetalFeat.ModifyDefinition SheetMetal, CurrentModel, Nothing
    
End Sub
