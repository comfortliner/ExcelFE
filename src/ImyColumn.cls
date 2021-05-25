VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ImyColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Model.Column")
Option Explicit

Public Property Get BgColor() As RenderingColor
End Property

Public Property Get CalculationField() As String
End Property

Public Property Get CharColor() As RenderingColor
End Property

Public Property Get columnChar() As String
End Property

Public Property Get PDirection() As ParameterDirectionEnum
End Property

Public Property Get ErrorMessage() As String
End Property

Public Property Get Formula() As String
End Property

Public Property Get Formula1() As String
End Property

Public Property Get Formula2() As String
End Property

Public Property Get FType() As FormattingType
End Property

Public Property Get Header() As String
End Property

Public Property Get HeaderColor() As RenderingColor
End Property

Public Property Get HeaderColorTaS() As Double
End Property

Public Property Get HeaderFontColor() As RenderingColor
End Property

Public Property Get IsGroup() As RenderingGrouping
End Property

Public Property Get IsReadOnly() As RenderingReadOnly
End Property

Public Property Get IsVisible() As RenderingVisible
End Property

Public Property Get Max() As Long
End Property

Public Property Get Min() As Long
End Property

Public Property Get PName() As String
End Property

Public Property Get Operator() As FormattingOperator
End Property

Public Property Get OtherColumn() As String
End Property

Public Property Get PSize() As Long
End Property

Public Property Get PType() As DataTypeEnum
End Property

Public Property Get VType() As ValidationType
End Property

Public Property Get Width() As Long
End Property

Public Property Get Zahl() As Long
End Property

Public Sub addFormatting( _
       ByVal FType As FormattingType, _
       Optional ByVal CharColor As RenderingColor = vbBlack, _
       Optional ByVal BgColor As RenderingColor = NoColor, _
       Optional ByVal Operator As FormattingOperator = NoOperator, _
       Optional ByVal OtherColumn As String = vbNullString, _
       Optional ByVal Zahl As Long = 0, _
       Optional ByVal Min As Long = 0, _
       Optional ByVal Max As Long = 0 _
       )
End Sub

Public Sub addFormatting2( _
       ByVal FType As FormattingType, _
       ByVal columnChar As String, _
       ByVal Formula As String, _
       Optional ByVal CharColor As RenderingColor = vbBlack, _
       Optional ByVal CharColorTaS As Double = 0, _
       Optional ByVal BgColor As RenderingColor = NoColor, _
       Optional ByVal BgColorTaS As Double = 0 _
       )
End Sub

Public Sub addValidation( _
       ByVal VType As ValidationType, _
       Optional ByVal Formula1 As String = vbNullString, _
       Optional ByVal Formula2 As String = vbNullString, _
       Optional ByVal ErrorMessage As String = vbNullString _
       )
End Sub

Public Sub addParameter( _
       ByVal PType As DataTypeEnum, _
       ByVal name As String, _
       ByVal Direction As ParameterDirectionEnum, _
       Optional ByVal Size As Long = 0 _
       )

End Sub

