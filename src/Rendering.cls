VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Rendering"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Controller.Rendering")
Option Explicit

'TODO ActiveWindow.zoom = zoom
'TODO  Range(RNG_RO_START).Select : ActiveWindow.FreezePanes = True
'TODO Outlook

Public Sub setLayout(ByVal column As ImyColumn)
  On Error GoTo Err_setLayout
  Dim rng As Range

  ' Ueberschrift und Farbgebung
  Set rng = shMain.Range(column.columnChar & conf_Props.ROWHEADER)
  rng.FormulaR1C1 = column.Header
  rng.Font.Bold = True
  rng.Interior.Pattern = xlSolid
  rng.Interior.PatternColorIndex = xlAutomatic
  rng.Interior.Color = column.HeaderColor
  rng.Interior.TintAndShade = column.HeaderColorTaS
  rng.Font.Color = column.HeaderFontColor
    
  ' IsGroup:
  ' Unterstreichen wenn Teil einer Gruppierung
  rng.Font.Underline = CBool(column.IsGroup)

  ' CalculationField:
  ' Kalkulationszelle setzen
  Set rng = shMain.Range(column.columnChar & conf_Props.ROWHEADER - 1)
  If (column.CalculationField <> vbNullString) Then
    rng.FormulaR1C1 = column.CalculationField
    rng.Interior.Color = vbYellow
    rng.Interior.TintAndShade = 0.4
  Else
    rng.Clear
  End If
    
  ' IsReadOnly:
  ' Darunter liegende Zellen grau faerben bei ReadOnly
  ' Bearbeitung freigeben wenn nicht ReadOnly
  Set rng = shMain.Range(column.columnChar & conf_Props.RSTART & ":" & column.columnChar & conf_Props.REND)
  If (column.IsReadOnly) Then
    rng.Interior.Color = vbWhite
    rng.Interior.TintAndShade = -0.1
    rng.Locked = True
  Else
    rng.ClearFormats
    rng.Locked = False
  End If

  ' IsVisible:
  ' Spalte ausblenden oder Spaltenbreite setzen
  Set rng = shMain.Range(column.columnChar & ":" & column.columnChar)
  If (column.IsVisible) Then
    rng.ColumnWidth = column.Width
  Else
    rng.ColumnWidth = 0
  End If
            
  Exit Sub
    
Err_setLayout:
  Err.Raise 10000, "Rendering.setLayout Column " & column.columnChar, Err.Description
  Err.Clear
End Sub

Public Sub setFormatting(ByVal column As ImyColumn)
  On Error GoTo Err_setFormatting

  Select Case column.FType
  Case NoFormatting
    Exit Sub
    
  Case Col_Op_OtherCol
    setCol_Op_OtherCol column
    
  Case Col_Op_Today
    setCol_Op_Today column
    
  Case Col_Op_Integer
    setCol_Op_Integer column
    
  Case Col_BarChart
    setCol_BarChart column

  End Select
  
  Exit Sub
  
Err_setFormatting:
  Err.Raise 10000, "Rendering.setFormatting Column " & column.columnChar, Err.Description
  Err.Clear
End Sub

Public Sub setFormatting2(ByVal column As ImyColumn)
  On Error GoTo Err_setFormatting2

  Select Case column.FType
  Case NoFormatting
    Exit Sub
  Case With_Formula
    setFormatting_with_Formula column
  End Select
  
  Exit Sub
  
Err_setFormatting2:
  Err.Raise 10000, "Rendering.setFormatting2 Column " & column.columnChar, Err.Description
  Err.Clear
End Sub

Public Sub setValidation(ByVal column As ImyColumn)
  On Error GoTo Err_setValidation

  Dim rng As Range
   
  Set rng = shMain.Range(column.columnChar & conf_Props.RSTART & ":" & column.columnChar & conf_Props.REND)
  rng.Validation.Delete
    
  Select Case column.VType
  Case NoValidation
    Exit Sub
    
  Case Numbers
    setNumbers column
    
  Case SingleChars
    setSingleChars column
    
  Case Dates
    setDates column
    
  Case DropDown
    setDropDown column
  
  End Select
    
  Exit Sub
  
Err_setValidation:
  Err.Raise 10000, "Rendering.setValidation Column " & column.columnChar, Err.Description
  Err.Clear
End Sub

'@Ignore ParameterCanBeByVal
Public Sub setFormCond(ByRef NewFormCond As ImyFormCond)
  On Error GoTo Err_setFormCond

  Select Case NewFormCond.FType
  Case NoFormatting
    Exit Sub
  Case With_Formula
    setFC_with_Formula NewFormCond
  Case Col_BarChart
    setFC_with_BarChart NewFormCond
  End Select
  
  Exit Sub


Err_setFormCond:
  Err.Raise 10000, "Rendering.setFormCond FormCond ", Err.Description
  Err.Clear
End Sub

'@Ignore ParameterCanBeByVal
Public Sub setSlicer(ByRef NewSlicer As ImySlicer)
  On Error GoTo Err_setSlicer

  Dim a As Long: a = NewSlicer.StepSize
  Dim x As Long: x = NewSlicer.StepsInX
  Dim y As Long: y = NewSlicer.StepsInY

  Dim offsetX As Long: offsetX = 10

  Dim slc As Variant
  
  Set slc = ActiveWorkbook.SlicerCaches.Add2(shMain.ListObjects(NewSlicer.Source), NewSlicer.SourceColumn). _
    slicers.Add( _
    shMain, , , _
    NewSlicer.SliderHeader, _
    NewSlicer.StartTileY * a + (NewSlicer.StartTileY - 1) * (a * y), _
    NewSlicer.StartTileX * a + (NewSlicer.StartTileX - 1) * (a * x) + offsetX, _
                                                                    (NewSlicer.StepTilesX - 1) * a + NewSlicer.StepTilesX * (a * x), _
                                                                    (NewSlicer.StepTilesY - 1) * a + NewSlicer.StepTilesY * (a * y) _
                                                                    )
            
  With slc
    .NumberOfColumns = NewSlicer.NumOfCols
    .DisplayHeader = True
    .DisableMoveResizeUI = True
    .SlicerCache.CrossFilterType = xlSlicerCrossFilterShowItemsWithNoData
    .SlicerCache.SortItems = xlSlicerSortAscending
    .SlicerCache.SortUsingCustomLists = True
    .ROWHEIGHT = 15
  End With
  
  ActiveWorkbook.SlicerCaches.Item("Datenschnitt_" & NewSlicer.SourceColumn).ClearManualFilter
  
  ActiveSheet.shapes.Range(Array(NewSlicer.SourceColumn)).Select
  Selection.Locked = msoFalse
  
  NewSlicer.Height = _
                   ActiveWorkbook.SlicerCaches.Item("Datenschnitt_" & NewSlicer.SourceColumn).slicers.Item(NewSlicer.SourceColumn).Height
  
  Exit Sub
  
Err_setSlicer:
  Err.Raise 10000, "Rendering.setSlicer Slicer " & NewSlicer.SourceColumn, Err.Description
  Err.Clear
End Sub

Public Sub setButton(ByVal NewButton As ImyButton, ByVal StartTileY As Double)
  On Error GoTo Err_setButton

  Dim a As Long: a = NewButton.StepSize
  Dim x As Long: x = NewButton.StepsInX
  
  Dim offsetX As Long: offsetX = 10
  
  ActiveSheet.shapes.AddShape(msoShapeRoundedRectangle, _
                              NewButton.StartTileX * a + (NewButton.StartTileX - 1) * (a * x) + offsetX, _
                              StartTileY, _
                              (NewButton.StepTilesX - 1) * a + NewButton.StepTilesX * (a * x), _
                              22 _
                              ).Select

  Selection.ShapeRange.ShapeStyle = msoShapeStylePreset27
  Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = NewButton.Caption
  Selection.ShapeRange.TextFrame2.VerticalAnchor = msoAnchorMiddle
  Selection.OnAction = NewButton.Command
    
  Exit Sub
  
Err_setButton:
  Err.Raise 10000, "Rendering.setButton Button " & NewButton.Caption, Err.Description
  Err.Clear
End Sub

Private Sub setFC_with_Formula(ByVal FC As ImyFormCond)
  On Error GoTo Err_setFC_with_Formula
  
  Dim rng As Range
  
  Set rng = shMain.Range(FC.FColumnChar & conf_Props.RSTART & ":" & FC.FColumnChar & conf_Props.REND)
  rng.FormatConditions.Add Type:=xlExpression, _
    Formula1:=FC.FFormula

  rng.FormatConditions.Item(rng.FormatConditions.Count).SetFirstPriority
  
  With rng.FormatConditions.Item(1).Font
    ' .Bold = True
    .Color = FC.FCharColor
    .TintAndShade = FC.FCharColorTaS
  End With
  
  If Not (FC.FBgColor = NoColor) Then
    With rng.FormatConditions.Item(1).Interior
      .Color = FC.FBgColor
      .TintAndShade = FC.FBgColorTaS
    End With
  End If
  
  rng.FormatConditions.Item(1).StopIfTrue = False
    
  Exit Sub

Err_setFC_with_Formula:
  Err.Raise 10000, "Rendering.setFormCond FC_with_Formula Column " & FC.FColumnChar, Err.Description
  Err.Clear
End Sub

Private Sub setFC_with_BarChart(ByVal FC As ImyFormCond)
  On Error GoTo Err_setFC_with_BarChart

  Dim rng As Range

  Set rng = shMain.Range(FC.FColumnChar & conf_Props.RSTART & ":" & FC.FColumnChar & conf_Props.REND)

  rng.FormatConditions.AddDatabar
  rng.FormatConditions.Item(rng.FormatConditions.Count).ShowValue = True
  rng.FormatConditions.Item(rng.FormatConditions.Count).SetFirstPriority
  With rng.FormatConditions.Item(1)
    .MinPoint.Modify newtype:=xlConditionValueNumber, newvalue:=FC.FMin
    .MaxPoint.Modify newtype:=xlConditionValueNumber, newvalue:=FC.FMax
  End With
  With rng.FormatConditions.Item(1).BarColor
    .Color = FC.FCharColor
    .TintAndShade = FC.FCharColorTaS
  End With
  rng.FormatConditions.Item(1).BarFillType = xlDataBarFillSolid
  rng.FormatConditions.Item(1).Direction = xlContext
  rng.FormatConditions.Item(1).NegativeBarFormat.ColorType = xlDataBarColor
  rng.FormatConditions.Item(1).BarBorder.Type = xlDataBarBorderNone
  rng.FormatConditions.Item(1).AxisPosition = xlDataBarAxisAutomatic
  With rng.FormatConditions.Item(1).AxisColor
    .Color = 0
    .TintAndShade = 0
  End With
  With rng.FormatConditions.Item(1).NegativeBarFormat.Color
    .Color = 255
    .TintAndShade = 0
  End With

  Exit Sub
  
Err_setFC_with_BarChart:
  Err.Raise 10000, "Rendering.setFormCond FC_with_BarChart Column " & FC.FColumnChar, Err.Description
  Err.Clear
End Sub


Private Sub setFormatting_with_Formula(ByVal column As ImyColumn)
  On Error GoTo Err_setFormatting_with_Formula
  
  Dim rng As Range
  
  Set rng = shMain.Range(column.columnChar & conf_Props.RSTART & ":" & column.columnChar & conf_Props.REND)
  rng.FormatConditions.Add Type:=xlExpression, _
    Formula1:=column.Formula

  rng.FormatConditions.Item(rng.FormatConditions.Count).SetFirstPriority
  
  With rng.FormatConditions.Item(1).Font
    .Bold = True
    .Color = column.CharColor
    .TintAndShade = 0
  End With
  
  rng.FormatConditions.Item(1).StopIfTrue = False
    
  Exit Sub

Err_setFormatting_with_Formula:
  Err.Raise 10000, "Rendering.setFormatting Formatting_with_Formula Column " & column.columnChar, Err.Description
  Err.Clear
End Sub

Private Sub setCol_Op_OtherCol(ByVal column As ImyColumn)
  On Error GoTo Err_Col_Op_OtherCol

  Dim rng As Range
  Dim Operator As String
  
  Select Case column.Operator
  Case Equal
    Operator = "="
  
  Case Greater
    Operator = ">"
  
  Case Less
    Operator = "<"
  
  Case Greater_Equal
    Operator = ">="
  
  Case Less_Equal
    Operator = "<="
  End Select
  
  Set rng = shMain.Range(column.columnChar & conf_Props.RSTART & ":" & column.columnChar & conf_Props.REND)

  rng.FormatConditions.Add Type:=xlExpression, _
                           Formula1:="=" & _
                                      column.columnChar & _
                                      conf_Props.RSTART & " " & Operator & " " & column.OtherColumn & conf_Props.RSTART
  rng.FormatConditions.Item(rng.FormatConditions.Count).SetFirstPriority
  With rng.FormatConditions.Item(1).Font
    .Bold = True
    .Color = column.CharColor
    .TintAndShade = 0
  End With
  If Not (column.BgColor = NoColor) Then
    With rng.FormatConditions.Item(1).Interior
      .Color = column.BgColor
    End With
  End If
  rng.FormatConditions.Item(1).StopIfTrue = False
    
  Exit Sub
  
Err_Col_Op_OtherCol:
  Err.Raise 10000, "Rendering.setFormatting Col_Op_OtherCol Column " & column.columnChar, Err.Description
  Err.Clear
End Sub

Private Sub setCol_Op_Today(ByVal column As ImyColumn)
  On Error GoTo Err_Col_Op_Today

  Dim rng As Range

  Set rng = shMain.Range(column.columnChar & conf_Props.RSTART & ":" & column.columnChar & conf_Props.REND)
  
  rng.FormatConditions.Add Type:=xlCellValue, Operator:=column.Operator, Formula1:="=HEUTE()"
  rng.FormatConditions.Item(rng.FormatConditions.Count).SetFirstPriority
  With rng.FormatConditions.Item(1).Font
    .Bold = True
    .Color = column.CharColor
  End With
  If Not (column.BgColor = NoColor) Then
    With rng.FormatConditions.Item(1).Interior
      .Color = column.BgColor
    End With
  End If
  rng.FormatConditions.Item(1).StopIfTrue = False
  Exit Sub
  
Err_Col_Op_Today:
  Err.Raise 10000, "Rendering.setFormatting Col_Op_Today Column " & column.columnChar, Err.Description
  Err.Clear
End Sub

Private Sub setCol_Op_Integer(ByVal column As ImyColumn)
  On Error GoTo Err_Col_Op_Integer

  Dim rng As Range
  Dim Operator As String
  
  Select Case column.Operator
  Case Equal
    Operator = "="

  Case Greater
    Operator = ">"

  Case Less
    Operator = "<"

  Case Greater_Equal
    Operator = ">="

  Case Less_Equal
    Operator = "<="
  End Select
  
  Set rng = shMain.Range(column.columnChar & conf_Props.RSTART & ":" & column.columnChar & conf_Props.REND)

  rng.FormatConditions.Add Type:=xlExpression, Formula1:="=" & column.columnChar & conf_Props.RSTART & " " & Operator & " " & column.Zahl
  rng.FormatConditions.Item(rng.FormatConditions.Count).SetFirstPriority
  With rng.FormatConditions.Item(1).Font
    .Bold = True
    .Color = column.CharColor
  End With
  If Not (column.BgColor = NoColor) Then
    With rng.FormatConditions.Item(1).Interior
      .Color = column.BgColor
    End With
  End If
  rng.FormatConditions.Item(1).StopIfTrue = False

  Exit Sub

Err_Col_Op_Integer:
  Err.Raise 10000, "Rendering.setFormatting Col_Op_Integer Column " & column.columnChar, Err.Description
  Err.Clear
End Sub

Private Sub setCol_BarChart(ByVal column As ImyColumn)
  On Error GoTo Err_Col_BarChart

  Dim rng As Range

  Set rng = shMain.Range(column.columnChar & conf_Props.RSTART & ":" & column.columnChar & conf_Props.REND)

  rng.FormatConditions.AddDatabar
  rng.FormatConditions.Item(rng.FormatConditions.Count).ShowValue = True
  rng.FormatConditions.Item(rng.FormatConditions.Count).SetFirstPriority
  With rng.FormatConditions.Item(1)
    .MinPoint.Modify newtype:=xlConditionValueNumber, newvalue:=column.Min
    .MaxPoint.Modify newtype:=xlConditionValueNumber, newvalue:=column.Max
  End With
  With rng.FormatConditions.Item(1).BarColor
    .Color = column.CharColor
    .TintAndShade = 0.65
  End With
  rng.FormatConditions.Item(1).BarFillType = xlDataBarFillSolid
  rng.FormatConditions.Item(1).Direction = xlContext
  rng.FormatConditions.Item(1).NegativeBarFormat.ColorType = xlDataBarColor
  rng.FormatConditions.Item(1).BarBorder.Type = xlDataBarBorderNone
  rng.FormatConditions.Item(1).AxisPosition = xlDataBarAxisAutomatic
  With rng.FormatConditions.Item(1).AxisColor
    .Color = 0
    .TintAndShade = 0
  End With
  With rng.FormatConditions.Item(1).NegativeBarFormat.Color
    .Color = 255
    .TintAndShade = 0
  End With

  Exit Sub

Err_Col_BarChart:
  Err.Raise 10000, "Rendering.setFormatting Col_BarChart Column " & column.columnChar, Err.Description
  Err.Clear
End Sub

Private Sub setNumbers(ByVal column As ImyColumn)
  On Error GoTo Err_setNumbers

  Dim rng As Range
   
  Set rng = shMain.Range(column.columnChar & conf_Props.RSTART & ":" & column.columnChar & conf_Props.REND)
    
  With rng.Validation
    .Delete
    .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
         Operator:=xlBetween, Formula1:=column.Formula1, Formula2:=column.Formula2
    .IgnoreBlank = True
    .InCellDropdown = False
    .InputTitle = vbNullString
    .ErrorTitle = "Fehler bei der Datenueberpruefung"
    .InputMessage = vbNullString
    .ErrorMessage = column.ErrorMessage
    .ShowInput = True
    .ShowError = True
  End With

  Exit Sub
  
Err_setNumbers:
  Err.Raise 10000, "Rendering.setValidation Numbers Column " & column.columnChar, Err.Description
  Err.Clear
End Sub

Private Sub setSingleChars(ByVal column As ImyColumn)
  On Error GoTo Err_setSingleChars

  Dim rng As Range
   
  Set rng = shMain.Range(column.columnChar & conf_Props.RSTART & ":" & column.columnChar & conf_Props.REND)

  With rng.Validation
    .Delete
    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
         Operator:=xlBetween, Formula1:=column.Formula1
    .IgnoreBlank = True
    .InCellDropdown = False
    .InputTitle = vbNullString
    .ErrorTitle = "Fehler bei der Datenueberpruefung"
    .InputMessage = vbNullString
    .ErrorMessage = column.ErrorMessage
    .ShowInput = True
    .ShowError = True
  End With

  Exit Sub
  
Err_setSingleChars:
  Err.Raise 10000, "Rendering.setValidation SingleChars Column " & column.columnChar, Err.Description
  Err.Clear
End Sub

Private Sub setDates(ByVal column As ImyColumn)
  On Error GoTo Err_setDates

  Dim rng As Range
   
  Set rng = shMain.Range(column.columnChar & conf_Props.RSTART & ":" & column.columnChar & conf_Props.REND)
    
  rng.NumberFormat = "m/d/yyyy"

  With rng.Validation
    .Delete
    .Add Type:=xlValidateDate, AlertStyle:=xlValidAlertStop, _
         Operator:=xlBetween, Formula1:=column.Formula1, Formula2:=column.Formula2
    .IgnoreBlank = True
    .InCellDropdown = False
    .InputTitle = vbNullString
    .ErrorTitle = "Fehler bei der Datenueberpruefung"
    .InputMessage = vbNullString
    .ErrorMessage = column.ErrorMessage
    .ShowInput = True
    .ShowError = True
  End With

  Exit Sub
  
Err_setDates:
  Err.Raise 10000, "Rendering.setValidation Dates Column " & column.columnChar, Err.Description
  Err.Clear
End Sub

Private Sub setDropDown(ByVal column As ImyColumn)
  On Error GoTo Err_setDropDown

  Dim rng As Range
 
  Set rng = shMain.Range(column.columnChar & conf_Props.RSTART & ":" & column.columnChar & conf_Props.REND)
    
  With rng.Validation
    .Delete
    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
         Operator:=xlBetween, Formula1:="=" & column.Formula1
    .IgnoreBlank = True
    .InCellDropdown = True
    .InputTitle = vbNullString
    .ErrorTitle = "Fehler bei der Datenueberpruefung"
    .InputMessage = vbNullString
    .ErrorMessage = column.ErrorMessage
    .ShowInput = True
    .ShowError = True
  End With

  Exit Sub
  
Err_setDropDown:
  Err.Raise 10000, "Rendering.setValidation DropDown Column " & column.columnChar, Err.Description
  Err.Clear
End Sub


