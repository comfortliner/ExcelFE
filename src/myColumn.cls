VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "myColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Model.Column")
Option Explicit

Implements ImyColumn

Private Type TColumn
  RCalculationField As String
  RColumnChar As String
  RDatabaseName As String
  RHEADER As String
  RHeaderColor As RenderingColor
  RHeaderColorTaS As Double
  RHeaderFontColor As RenderingColor
  RIsGroup As RenderingGrouping
  RIsReadOnly As RenderingReadOnly
  RIsVisible As RenderingVisible
  RWidth As Long

  FBgColor As RenderingColor
  FBgColorTaS As Double
  FColumnChar As String
  FCharColor As RenderingColor
  FCharColorTaS As Double
  FFormula As String
  FMax As Long
  FMin As Long
  FOperator As FormattingOperator
  FOtherColumn As String
  FType As FormattingType
  FZahl As Long
  
  VErrorMessage As String
  VFormula1 As String
  VFormula2 As String
  VType As ValidationType
  
  PName As String
  PType As DataTypeEnum
  PDirection As ParameterDirectionEnum
  PSize As Long
End Type

Private This As TColumn

Private Property Get ImyColumn_CalculationField() As String
  ImyColumn_CalculationField = This.RCalculationField
End Property

Private Property Get ImyColumn_ColumnChar() As String
  ImyColumn_ColumnChar = This.RColumnChar
End Property

Private Property Get ImyColumn_PDirection() As ParameterDirectionEnum
  ImyColumn_PDirection = This.PDirection
End Property

Private Property Get ImyColumn_ErrorMessage() As String
  ImyColumn_ErrorMessage = This.VErrorMessage
End Property

Private Property Get ImyColumn_Formula1() As String
  ImyColumn_Formula1 = This.VFormula1
End Property

Private Property Get ImyColumn_Formula2() As String
  ImyColumn_Formula2 = This.VFormula2
End Property

Private Property Get ImyColumn_FType() As FormattingType
  ImyColumn_FType = This.FType
End Property

Private Property Get ImyColumn_Header() As String
  ImyColumn_Header = This.RHEADER
End Property

Private Property Get ImyColumn_HeaderColor() As RenderingColor
  ImyColumn_HeaderColor = This.RHeaderColor
End Property

Private Property Get ImyColumn_HeaderColorTaS() As Double
  ImyColumn_HeaderColorTaS = This.RHeaderColorTaS
End Property

Private Property Get ImyColumn_HeaderFontColor() As RenderingColor
  ImyColumn_HeaderFontColor = This.RHeaderFontColor
End Property

Private Property Get ImyColumn_IsGroup() As RenderingGrouping
  ImyColumn_IsGroup = This.RIsGroup
End Property

Private Property Get ImyColumn_IsReadOnly() As RenderingReadOnly
  ImyColumn_IsReadOnly = This.RIsReadOnly
End Property

Private Property Get ImyColumn_IsVisible() As RenderingVisible
  ImyColumn_IsVisible = This.RIsVisible
End Property

Private Property Get ImyColumn_CharColor() As RenderingColor
  ImyColumn_CharColor = This.FCharColor
End Property

Private Property Get ImyColumn_BgColor() As RenderingColor
  ImyColumn_BgColor = This.FBgColor
End Property

Private Property Get ImyColumn_Formula() As String
  ImyColumn_Formula = This.FFormula
End Property

Private Property Get ImyColumn_Max() As Long
  ImyColumn_Max = This.FMax
End Property

Private Property Get ImyColumn_Min() As Long
  ImyColumn_Min = This.FMin
End Property

Private Property Get ImyColumn_PName() As String
  ImyColumn_PName = This.PName
End Property

Private Property Get ImyColumn_Operator() As FormattingOperator
  ImyColumn_Operator = This.FOperator
End Property

Private Property Get ImyColumn_OtherColumn() As String
  ImyColumn_OtherColumn = This.FOtherColumn
End Property

Private Property Get ImyColumn_PSize() As Long
  ImyColumn_PSize = This.PSize
End Property

Private Property Get ImyColumn_PType() As DataTypeEnum
  ImyColumn_PType = This.PType
End Property

Private Property Get ImyColumn_VType() As ValidationType
  ImyColumn_VType = This.VType
End Property

Private Property Get ImyColumn_Width() As Long
  ImyColumn_Width = This.RWidth
End Property

Private Property Get ImyColumn_Zahl() As Long
  ImyColumn_Zahl = This.FZahl
End Property

Public Sub FillData( _
       ByVal columnChar As String, _
       ByVal Header As String, _
       ByVal DatabaseName As String, _
       Optional ByVal HeaderColor As RenderingColor = Green, _
       Optional ByVal HeaderColorTaS As Double = 0.3, _
       Optional ByVal HeaderFontColor As RenderingColor = Black, _
       Optional ByVal Width As Long = 10.78, _
       Optional ByVal CalculationField As String = vbNullString, _
       Optional ByVal IsGroup As RenderingGrouping = NoGroup, _
       Optional ByVal IsReadOnly As RenderingReadOnly = RO, _
       Optional ByVal IsVisible As RenderingVisible = Visible _
       )
  On Error GoTo Err_FillData

  If HeaderColorTaS < -1 Or HeaderColorTaS > 1 Then
    Err.Raise vbObjectError + 513, "Column " & columnChar & " .FillData", "HeaderColorTaS muss im Bereich von -1.0 und +1.0 liegen."
  End If
  
  If Len(This.RColumnChar) = 0 Then
    This.RCalculationField = CalculationField
    This.RColumnChar = columnChar
    This.RDatabaseName = DatabaseName
    This.RHEADER = Header
    This.RHeaderColor = HeaderColor
    This.RHeaderColorTaS = HeaderColorTaS
    This.RHeaderFontColor = HeaderFontColor
    This.RIsGroup = IsGroup
    This.RIsReadOnly = IsReadOnly
    This.RIsVisible = IsVisible
    This.RWidth = Width

    This.FBgColor = Black
    This.FCharColor = Black
    This.FFormula = vbNullString
    This.FMax = 0
    This.FMin = 0
    This.FOperator = NoOperator
    This.FOtherColumn = vbNullString
    This.FType = NoFormatting
    This.FZahl = 0
    
    This.VErrorMessage = vbNullString
    This.VFormula1 = vbNullString
    This.VFormula2 = vbNullString
    This.VType = NoValidation
  End If

  Exit Sub
  
Err_FillData:
  myLogger.LogError Err.Number & ": " & Err.Source & " / " & Err.Description, "myColumn.FillData"
  Err.Clear
End Sub

Private Sub ImyColumn_addFormatting( _
        ByVal FType As FormattingType, _
        Optional ByVal CharColor As RenderingColor = vbBlack, _
        Optional ByVal BgColor As RenderingColor = NoColor, _
        Optional ByVal Operator As FormattingOperator = NoOperator, _
        Optional ByVal OtherColumn As String = vbNullString, _
        Optional ByVal Zahl As Long = 0, _
        Optional ByVal Min As Long = 0, _
        Optional ByVal Max As Long = 0 _
        )
  On Error GoTo Err_addFormatting

  Select Case FType
  Case NoFormatting
    Exit Sub
    
  Case Col_Op_OtherCol
    If Operator = NoOperator Or OtherColumn = vbNullString Then
      Err.Raise vbObjectError + 513, "Column " & This.RColumnChar & " Operator oder OtherColumn fehlt"
    End If
    
  Case Col_Op_Today, Col_Op_Integer
    If Operator = NoOperator Then
      Err.Raise vbObjectError + 513, "Column " & This.RColumnChar & " Operator fehlt"
    End If

  End Select

  This.FType = FType
  This.FCharColor = CharColor
  This.FBgColor = BgColor
  This.FOperator = Operator
  This.FOtherColumn = OtherColumn
  This.FZahl = Zahl
  This.FMin = Min
  This.FMax = Max
  
  Exit Sub
  
Err_addFormatting:
  myLogger.LogError Err.Number & ": " & Err.Source & " / " & Err.Description, "myColumn.addFormatting"
  Err.Clear
End Sub

Private Sub ImyColumn_addFormatting2( _
       ByVal FType As FormattingType, _
       ByVal columnChar As String, _
       ByVal Formula As String, _
       Optional ByVal CharColor As RenderingColor = vbBlack, _
       Optional ByVal CharColorTaS As Double = 0, _
       Optional ByVal BgColor As RenderingColor = NoColor, _
       Optional ByVal BgColorTaS As Double = 0 _
       )
  On Error GoTo Err_addFormatting2
  
  Select Case FType
  Case NoFormatting
    Exit Sub
    
  End Select
  
  If columnChar = vbNullString Or Formula = vbNullString Then
    Err.Raise vbObjectError + 513, "columnChar oder Formula fehlen"
  End If
    
  This.FType = FType
  This.FBgColor = BgColor
  This.FBgColorTaS = BgColorTaS
  This.FColumnChar = columnChar
  This.FCharColor = CharColor
  This.FCharColorTaS = CharColorTaS
  This.FFormula = Formula
  
  Exit Sub
  
Err_addFormatting2:
  myLogger.LogError Err.Number & ": " & Err.Source & " / " & Err.Description, "myColumn.addFormatting2"
  Err.Clear

End Sub

Private Sub ImyColumn_addValidation( _
        ByVal VType As ValidationType, _
        Optional ByVal Formula1 As String = vbNullString, _
        Optional ByVal Formula2 As String = vbNullString, _
        Optional ByVal ErrorMessage As String = vbNullString _
        )
  On Error GoTo Err_addValidation
  
  Select Case VType
  Case NoValidation
    Exit Sub
    
  Case Numbers, Dates
    If Formula1 = vbNullString Or Formula2 = vbNullString Then
      Err.Raise vbObjectError + 513, "Column " & This.RColumnChar & " Formula1 oder Formula2 fehlt"
    End If
    
  Case SingleChars, DropDown
    If Formula1 = vbNullString Then
      Err.Raise vbObjectError + 513, "Column " & This.RColumnChar & " Formula1 fehlt"
    End If
  End Select
  
  This.VType = VType
  This.VFormula1 = Formula1
  This.VFormula2 = Formula2
  This.VErrorMessage = ErrorMessage

  Exit Sub
  
Err_addValidation:
  myLogger.LogError Err.Number & ": " & Err.Source & " / " & Err.Description, "myColumn.addValidation"
  Err.Clear
End Sub

Private Sub ImyColumn_addParameter( _
       ByVal PType As DataTypeEnum, _
       ByVal name As String, _
       ByVal Direction As ParameterDirectionEnum, _
       Optional ByVal Size As Long = 0 _
       )
    On Error GoTo Err_addParameter
    
    This.PType = PType
    This.PName = name
    This.PDirection = Direction
    This.PSize = Size
    
    Exit Sub
    
Err_addParameter:
  myLogger.LogError Err.Number & ": " & Err.Source & " / " & Err.Description, "myColumn.addParameter"
  Err.Clear
End Sub

