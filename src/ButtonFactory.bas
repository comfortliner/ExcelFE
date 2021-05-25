Attribute VB_Name = "ButtonFactory"
'@Folder("Controller")
Option Explicit

Public Function Create( _
       ByVal StartTileX As Long, _
       ByVal StepTilesX As Long, _
       ByVal Caption As String, _
       Optional ByVal Command As String = "btn_emptyCommand", _
       Optional ByVal VisibleFor As String = vbNullString _
       ) As ImyButton
  
  Dim NewButton As myButton
  Set NewButton = New myButton
  
  NewButton.FillData StartTileX, StepTilesX, Caption, Command
  
  Set Create = NewButton
End Function

Public Sub deleteAllButtons()
  Dim shp As Shape

  For Each shp In ActiveSheet.shapes
    If shp.Type = 1 Then shp.Delete
  Next

End Sub

Public Sub btn_emptyCommand()
  MsgBox "Fuer diesen Button wurde noch keine Sub hinterlegt."
End Sub

Public Sub btn_clearFilter()
  SlicerFactory.clearFilter
End Sub

Public Sub btn_tableSync()
On Error GoTo Err_btn_tableSync
Dim rng As Range
  
  conf_Props.InitializationFinished = False
  
  Utils.setProtection shMain, False
  
  ActiveWorkbook.RefreshAll
  
  Set rng = shMain.Range(conf_Props.LOGOUTPUT).Offset(-1, 0)
  rng.FormulaR1C1 = "Zeitstempel Tabelle synchronisieren: " & Format$(CDate(Now), "dd.mm.yyyy   hh:mm") & " Uhr."
  

Exit_btn_tableSync:
  Utils.setProtection shMain, True
  conf_Props.InitializationFinished = True
Exit Sub

Err_btn_tableSync:
  myLogger.LogError Err.Number & ": " & Err.Source & " / " & Err.Description, "ButtonFactory.btn_tableSync"
  Resume Exit_btn_tableSync
End Sub


