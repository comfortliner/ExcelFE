VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "mySlicer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Model.Slicer")
Option Explicit

Implements ImySlicer

Private Type TSlicer
  StartTileX As Long
  StartTileY As Long
  StepTilesX As Long
  StepTilesY As Long
  Source As String
  SourceColumn As String
  SliderHeader As String
  NumOfCols As Long
  
  StepSize As Long
  StepsInX As Long
  StepsInY As Long
  
  Height As Double
End Type

Private This As TSlicer

Private Property Let ImySlicer_Height(ByVal value As Double)
  This.Height = value
End Property

Private Property Get ImySlicer_Height() As Double
  ImySlicer_Height = This.Height
End Property

Private Property Get ImySlicer_StartTileX() As Long
  ImySlicer_StartTileX = This.StartTileX
End Property

Private Property Get ImySlicer_StartTileY() As Long
  ImySlicer_StartTileY = This.StartTileY
End Property

Private Property Get ImySlicer_StepSize() As Long
  ImySlicer_StepSize = This.StepSize
End Property

Private Property Get ImySlicer_StepsInX() As Long
  ImySlicer_StepsInX = This.StepsInX
End Property

Private Property Get ImySlicer_StepsInY() As Long
  ImySlicer_StepsInY = This.StepsInY
End Property

Private Property Get ImySlicer_StepTilesX() As Long
  ImySlicer_StepTilesX = This.StepTilesX
End Property

Private Property Get ImySlicer_StepTilesY() As Long
  ImySlicer_StepTilesY = This.StepTilesY
End Property

Private Property Get ImySlicer_Source() As String
  ImySlicer_Source = This.Source
End Property

Private Property Get ImySlicer_SourceColumn() As String
  ImySlicer_SourceColumn = This.SourceColumn
End Property

Private Property Get ImySlicer_SliderHeader() As String
  ImySlicer_SliderHeader = This.SliderHeader
End Property

Private Property Get ImySlicer_NumOfCols() As Long
  ImySlicer_NumOfCols = This.NumOfCols
End Property

Public Sub FillData( _
       ByVal StartTileX As Long, _
       ByVal StartTileY As Long, _
       ByVal StepTilesX As Long, _
       ByVal StepTilesY As Long, _
       ByVal Source As String, _
       ByVal SourceColumn As String, _
       ByVal SliderHeader As String, _
       Optional ByVal NumOfCols As Long = 1 _
       )
  
  If Len(This.SourceColumn) = 0 Then
    This.StartTileX = StartTileX
    This.StartTileY = StartTileY
    This.StepTilesX = StepTilesX
    This.StepTilesY = StepTilesY
    This.Source = Source
    This.SourceColumn = SourceColumn
    This.SliderHeader = SliderHeader
    This.NumOfCols = NumOfCols
    
    This.StepSize = 5
    This.StepsInX = 9
    This.StepsInY = 9
    
    This.Height = 0
  End If

  Exit Sub

End Sub


