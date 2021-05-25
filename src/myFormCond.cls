VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "myFormCond"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Model.FormCond")
Option Explicit

Implements ImyFormCond

Private Type TFormCond
  FBgColor As RenderingColor
  FBgColorTaS As Double
  FCharColor As RenderingColor
  FCharColorTaS As Double
  FColumnChar As String
  FFormula As String
  FMax As Long
  FMin As Long
  FType As FormattingType
End Type

Private This As TFormCond

Private Property Get ImyFormCond_FBgColor() As RenderingColor
  ImyFormCond_FBgColor = This.FBgColor
End Property

Private Property Get ImyFormCond_FBgColorTaS() As Double
  ImyFormCond_FBgColorTaS = This.FBgColorTaS
End Property

Private Property Get ImyFormCond_FCharColor() As RenderingColor
  ImyFormCond_FCharColor = This.FCharColor
End Property

Private Property Get ImyFormCond_FCharColorTaS() As Double
  ImyFormCond_FCharColorTaS = This.FCharColorTaS
End Property

Private Property Get ImyFormCond_FColumnChar() As String
  ImyFormCond_FColumnChar = This.FColumnChar
End Property

Private Property Get ImyFormCond_FFormula() As String
  ImyFormCond_FFormula = This.FFormula
End Property

Private Property Get ImyFormCond_FMax() As Long
  ImyFormCond_FMax = This.FMax
End Property

Private Property Get ImyFormCond_FMin() As Long
  ImyFormCond_FMin = This.FMin
End Property

Private Property Get ImyFormCond_FType() As FormattingType
  ImyFormCond_FType = This.FType
End Property

Public Sub FillData( _
       ByVal FType As FormattingType, _
       ByVal FColumnChar As String, _
       ByVal FFormula As String, _
       Optional ByVal FCharColor As RenderingColor = vbBlack, _
       Optional ByVal FCharColorTaS As Double = 0, _
       Optional ByVal FBgColor As RenderingColor = NoColor, _
       Optional ByVal FBgColorTaS As Double = 0, _
       Optional ByVal FMin As Long = 0, _
       Optional ByVal FMax As Long = 100 _
)
  On Error GoTo Err_FillData

  If FCharColorTaS < -1 Or FCharColorTaS > 1 Then
    Err.Raise vbObjectError + 513, "Column: " & FColumnChar, "CharColorTaS muss im Bereich von -1.0 und +1.0 liegen."
  End If
  
  If FBgColorTaS < -1 Or FBgColorTaS > 1 Then
    Err.Raise vbObjectError + 513, "Column: " & FColumnChar, "BgColorTaS muss im Bereich von -1.0 und +1.0 liegen."
  End If
  
  If FColumnChar = vbNullString Then
    Err.Raise vbObjectError + 513, "Column: " & FColumnChar, "ColumnChar darf nicht leer sein."
  End If
  
  If FFormula = vbNullString Then
    Err.Raise vbObjectError + 513, "Column: " & FColumnChar, "Formula darf nicht leer sein."
  End If
  
  If Left$(FFormula, 1) <> "=" Then
    Err.Raise vbObjectError + 513, "Column: " & FColumnChar, "Formula muss mit '=' beginnen."
  End If
  
  This.FBgColor = FBgColor
  This.FBgColorTaS = FBgColorTaS
  This.FCharColor = FCharColor
  This.FCharColorTaS = FCharColorTaS
  This.FColumnChar = FColumnChar
  This.FFormula = FFormula
  This.FMax = FMax
  This.FMin = FMin
  This.FType = FType
  
  Exit Sub

Err_FillData:
  myLogger.LogError Err.Number & ": " & Err.Source & " / " & Err.Description, "myFormCond.FillData"
  Err.Clear
End Sub



