VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ImySlicer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Model.Slicer")
Option Explicit

Public Property Let Height(ByVal value As Double)
End Property

Public Property Get Height() As Double
End Property

Public Property Get StartTileX() As Long
End Property

Public Property Get StartTileY() As Long
End Property

Public Property Get StepSize() As Long
End Property

Public Property Get StepsInX() As Long
End Property

Public Property Get StepsInY() As Long
End Property

Public Property Get StepTilesX() As Long
End Property

Public Property Get StepTilesY() As Long
End Property

Public Property Get Source() As String
End Property

Public Property Get SourceColumn() As String
End Property

Public Property Get SliderHeader() As String
End Property

Public Property Get NumOfCols() As Long
End Property

