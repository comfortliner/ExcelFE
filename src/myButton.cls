VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "myButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Model.Button")
Option Explicit

Implements ImyButton

Private Type TButton
  StartTileX As Long
  StepTilesX As Long
  Caption As String
  Command As String
  VisibleFor As String
  
  StepSize As Long
  StepsInX As Long
End Type

Private This As TButton

Private Property Get ImyButton_StartTileX() As Long
  ImyButton_StartTileX = This.StartTileX
End Property

Private Property Get ImyButton_StepTilesX() As Long
  ImyButton_StepTilesX = This.StepTilesX
End Property

Private Property Get ImyButton_Caption() As String
  ImyButton_Caption = This.Caption
End Property

Private Property Get ImyButton_Command() As String
  ImyButton_Command = This.Command
End Property

Private Property Get ImyButton_VisibleFor() As String
  ImyButton_VisibleFor = This.VisibleFor
End Property

Private Property Get ImyButton_StepSize() As Long
  ImyButton_StepSize = This.StepSize
End Property

Private Property Get ImyButton_StepsInX() As Long
  ImyButton_StepsInX = This.StepsInX
End Property

Public Sub FillData( _
       ByVal StartTileX As Long, _
       ByVal StepTilesX As Long, _
       ByVal Caption As String, _
       Optional ByVal Command As String = "btn_emptyCommand", _
       Optional ByVal VisibleFor As String = vbNullString _
       )
  
  This.StartTileX = StartTileX
  This.StepTilesX = StepTilesX
  This.Caption = Caption
  This.Command = Command
  
  This.StepSize = 5
  This.StepsInX = 9

  Exit Sub
  
End Sub


