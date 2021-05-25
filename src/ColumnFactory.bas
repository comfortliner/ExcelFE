Attribute VB_Name = "ColumnFactory"
'@Folder("Controller")
Option Explicit

Public Function Create( _
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
       ) As ImyColumn
  
  Dim NewColumn As myColumn
  Set NewColumn = New myColumn
  
  NewColumn.FillData columnChar, Header, DatabaseName, HeaderColor, HeaderColorTaS, HeaderFontColor, _
                     Width, CalculationField, IsGroup, IsReadOnly, IsVisible
  Set Create = NewColumn
End Function


