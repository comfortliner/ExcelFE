Attribute VB_Name = "FormCondFactory"
'@Folder("Controller")
Option Explicit

Public Function Create( _
       ByVal FType As FormattingType, _
       ByVal FColumnChar As String, _
       ByVal FFormula As String, _
       Optional ByVal FCharColor As RenderingColor = vbBlack, _
       Optional ByVal FCharColorTaS As Double = 0, _
       Optional ByVal FBgColor As RenderingColor = NoColor, _
       Optional ByVal FBgColorTaS As Double = 0, _
       Optional ByVal FMin As Long = 0, _
       Optional ByVal FMax As Long = 100 _
) As ImyFormCond

  Dim NewFormCond As myFormCond
  Set NewFormCond = New myFormCond
  
  NewFormCond.FillData FType, FColumnChar, FFormula, FCharColor, FCharColorTaS, FBgColor, FBgColorTaS, _
                       FMin, FMax
                       
  Set Create = NewFormCond
End Function

