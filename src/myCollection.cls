VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "myCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Model.Collection")
Option Explicit

Implements ImyCollection

Private Type TCollection
  buttons As VBA.Collection
  columns As VBA.Collection
  slicers As VBA.Collection
  dbDefaultParams As VBA.Collection
  FormCond As VBA.Collection
End Type

Private This As TCollection

Private Sub Class_Initialize()
  Set This.buttons = New VBA.Collection
  Set This.columns = New VBA.Collection
  Set This.slicers = New VBA.Collection
  Set This.dbDefaultParams = New VBA.Collection
  Set This.FormCond = New VBA.Collection
End Sub

Private Sub Class_Terminate()
  Set This.buttons = Nothing
  Set This.columns = Nothing
  Set This.slicers = Nothing
  Set This.dbDefaultParams = Nothing
  Set This.FormCond = Nothing
End Sub

Private Property Get ImyCollection_columnsCollection() As VBA.Collection
  Set ImyCollection_columnsCollection = This.columns
End Property

Private Property Get ImyCollection_paramsCollection() As VBA.Collection
  Set ImyCollection_paramsCollection = This.dbDefaultParams
End Property

Private Property Get ImyCollection_formcondCollection() As VBA.Collection
  Set ImyCollection_formcondCollection = This.FormCond
End Property

'Public Sub FillData()
'End Sub

Private Sub ImyCollection_addColumn(ByVal column As ImyColumn)
  This.columns.Add column
End Sub

Private Sub ImyCollection_addSlicer(ByVal Slicer As ImySlicer)
  This.slicers.Add Slicer
End Sub

Private Sub ImyCollection_addButton(ByVal Button As ImyButton)
  This.buttons.Add Button
End Sub

Private Sub ImyCollection_adddbDefaultParam(ByVal Param As ImyDBDefaultParam)
  This.dbDefaultParams.Add Param
End Sub

Private Sub ImyCollection_addFormCond(ByVal FormCond As ImyFormCond)
  This.FormCond.Add FormCond
End Sub

Private Sub ImyCollection_generateshMain()
  generateColumns
End Sub

Private Sub ImyCollection_initializeshMain()
  generateFormCond
  generateSlicers
  generateButtons
End Sub

Private Sub generateColumns()
  On Error GoTo Err_generateColumns
  
  Dim column As ImyColumn
  Dim render As Rendering

  Set render = New Rendering
  Application.ScreenUpdating = False
    
  shMain.Cells.FormatConditions.Delete
  
  For Each column In This.columns
    render.setLayout column
    ' render.setFormatting column
    ' render.setFormatting2 column
    render.setValidation column
  Next column
  

  Application.ScreenUpdating = True
  Set render = Nothing
  Exit Sub
  
Err_generateColumns:
  myLogger.LogError Err.Number & ": " & Err.Source & " / " & Err.Description, "Collection.generateColumns"
  Err.Clear
  
  Application.ScreenUpdating = True
  Set render = Nothing
End Sub

Private Sub generateFormCond()
  On Error GoTo Err_generateFormCond
  
  Dim FormCond As ImyFormCond
  Dim render As Rendering

  Application.Cursor = xlWait
  Set render = New Rendering
  Application.ScreenUpdating = False
  
  shMain.Cells.FormatConditions.Delete
  
  ' Form Conditions erzeugen
  For Each FormCond In This.FormCond
    render.setFormCond FormCond
  Next FormCond


Exit_generateFormCond:
  Application.ScreenUpdating = True
  Set render = Nothing
  Application.Cursor = xlDefault
  Exit Sub
  
Err_generateFormCond:
  myLogger.LogError Err.Number & ": " & Err.Source & " / " & Err.Description, "Collection.generateFormCond"
  Err.Clear
  Resume Exit_generateFormCond
End Sub

Private Sub generateSlicers()
  On Error GoTo Err_generateSlicers
  
  Dim Slicer As ImySlicer
  Dim render As Rendering

  Application.Cursor = xlWait
  Set render = New Rendering
  Application.ScreenUpdating = False
  
  SlicerFactory.deleteAllSlicer
  
  ' Slicer erzeugen
  For Each Slicer In This.slicers
    render.setSlicer Slicer
  Next Slicer
  
  
  ' Zeile 1 als Hintergrund der Slicer formatieren
  With shMain.Range("1:1")
    .ROWHEIGHT = getMaxSlicerHeight
    .Interior.Color = vbWhite
    .Interior.TintAndShade = -0.2
  End With


Exit_generateSlicers:
  Application.ScreenUpdating = True
  Set render = Nothing
  Application.Cursor = xlDefault
  Exit Sub
  
Err_generateSlicers:
  myLogger.LogError Err.Number & ": " & Err.Source & " / " & Err.Description, "Collection.generateSlicers"
  Err.Clear
  Resume Exit_generateSlicers
End Sub

Private Sub generateButtons()
  On Error GoTo Err_generateButtons
  
  Dim Button As ImyButton
  Dim render As Rendering

  Application.Cursor = xlWait
  Set render = New Rendering
  Application.ScreenUpdating = False
  
  ButtonFactory.deleteAllButtons
  
  ' Button erzeugen
  For Each Button In This.buttons
    render.setButton Button, getMaxSlicerHeight + shMain.Range("2:2").Height + 4
  Next Button
  
  ' Zeile 3 als Hintergrund der Buttons formatieren
  With shMain.Range("3:3")
    .ROWHEIGHT = 30
    .Interior.Color = vbWhite
    .Interior.TintAndShade = -0.1
  End With

Exit_generateButtons:
  Application.ScreenUpdating = True
  Set render = Nothing
  Application.Cursor = xlDefault
  Exit Sub
  
Err_generateButtons:
  myLogger.LogError Err.Number & ": " & Err.Source & " / " & Err.Description, "Collection.generateButtons"
  Err.Clear
  Resume Exit_generateButtons
End Sub

Private Function getMaxSlicerHeight() As Double
  Dim Slicer As ImySlicer

  Dim resRowHeight As Double: resRowHeight = 14.4
  Dim resStepSize As Long: resStepSize = 0

  For Each Slicer In This.slicers
    If Slicer.Height > resRowHeight Then resRowHeight = Slicer.Height
    If Slicer.StepSize > resStepSize Then resStepSize = Slicer.StepSize
  Next Slicer

  getMaxSlicerHeight = resRowHeight + 2 * resStepSize

End Function

