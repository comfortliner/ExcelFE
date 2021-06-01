VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Utils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("Controller")
'@PredeclaredId
Option Explicit

Private Declare Function ShellExecute _
  Lib "shell32.dll" Alias "ShellExecuteA" ( _
  ByVal hWnd As Long, _
  ByVal Operation As String, _
  ByVal Filename As String, _
  Optional ByVal Parameters As String, _
  Optional ByVal Directory As String, _
  Optional ByVal WindowStyle As Long = vbMinimizedFocus _
  ) As Long
  
Private Declare Function URLDownloadToFile Lib "urlmon" Alias _
  "URLDownloadToFileA" ( _
  ByVal pCaller As Long, ByVal szURL As String, _
  ByVal szFileName As String, _
  ByVal dwReserved As Long, _
  ByVal lpfnCB As Long _
  ) As Long

Public Sub openURL(ByVal url As String)
On Error GoTo Err_OpenURL
Dim success As Long

  success = ShellExecute(0, "Open", url)

Exit_OpenURL:
  Exit Sub

Err_OpenURL:
  myLogger.LogError Err.Number & ": " & Err.Source & " / " & Err.Description, "Utils.openURL"
  Resume Exit_OpenURL

End Sub


Sub SaveFileFromSPInUserProfileAndOpenIt()
' Copy and use this Sub into Start_Plama...xlsb
On Error Resume Next

Dim ThisWorkbook As Workbook
Dim fso As New FileSystemObject
Dim fileToCopy As String: fileToCopy = Mid(ActiveWorkbook.name, InStr(ActiveWorkbook.name, "_") + 1, 30)
Dim SourceSPUrl As String
Dim UserProfile, DestinationPath, DestinationSubPath As String
Dim Source, Destination As String
Dim returnValue As Long

  
  SourceSPUrl = ActiveWorkbook.Path
  
  UserProfile = Environ("UserProfile")
  DestinationSubPath = "PlaMa"
  DestinationPath = UserProfile & "\" & DestinationSubPath
  
  Source = SourceSPUrl & "/_DO_NOT_USE/" & fileToCopy
  Destination = DestinationPath & "\" & fileToCopy
  
  
  
  Application.ScreenUpdating = False
  
  If Not fso.FolderExists(DestinationPath) Then
    fso.CreateFolder DestinationPath
  End If

  returnValue = URLDownloadToFile(0, Source, Destination, 0, 0)
    
  Application.ScreenUpdating = True
    
  If returnValue = 0 Then
    Set ThisWorkbook = ActiveWorkbook
    Workbooks.Open (Destination)
    ThisWorkbook.Close
  Else
    MsgBox "Fehler beim Kopieren der Datei " & fileToCopy & vbCrLf & _
      "Startdatei wird jetzt geschlossen."
    
    If Application.Workbooks.Count = 1 Then
      Application.Quit
    Else
      ActiveWorkbook.Close
    End If
  End If
  
End Sub

Public Sub writeToDB(ByVal Target As Range)
On Error GoTo Err_writeToDB

  Dim column1 As String
  Dim column2 As String
  Dim row1 As Long
  Dim row2 As Long

  With conf_Props
    column1 = .CRWSTART
    column2 = .CRWEND
    row1 = .RSTART
    row2 = .REND
  End With

  If (Not Application.Intersect(Target, ActiveSheet.Range(column1 & row1 & ":" & column2 & row2)) Is Nothing) And _
     (conf_Props.InitializationFinished) Then
    Dim db As myMSSQL
    Set db = New myMSSQL
    db.execute Target.row
  End If

Exit_writeToDB:
  If conf_Props.InitializationFinished Then _
     If Not db Is Nothing Then Set db = Nothing
  Exit Sub

Err_writeToDB:
  myLogger.LogError Err.Number & ": " & Err.Source & " / " & Err.Description, "Utils.writeToDB"
  Resume Exit_writeToDB
End Sub

Public Function showHairlineCross(ByVal Target As Range) As Boolean
' Blendet ein Fadenkreuz ein zur besseren Orientierung
Dim Selection As String

    If Target.Cells.column <= ActiveSheet.Range(conf_Props.CRWEND & conf_Props.RSTART).column Then
        showHairlineCross = False
        Exit Function
    End If
    
    Selection = Target.EntireColumn.Address & "," & Target.EntireRow.Address
    ActiveSheet.Range(Selection).Select
    showHairlineCross = True
End Function

Public Sub preventMultiSelection(ByVal Target As Range)
' Verhindert, dass mehrere Zellen auf einmal ausgewaehlt werden koennen
' Soll nur fuer die Spalte im SAP Rohdaten Read-Only Bereich moeglich sein
On Error GoTo Err_preventMultiSelection
    
    If Target.Cells.column <= ActiveSheet.Range(conf_Props.CRWEND & conf_Props.RSTART).column And conf_Props.InitializationFinished Then
       Application.EnableEvents = False
       ActiveCell.Select
    End If

Err_preventMultiSelection:
    Application.EnableEvents = True
End Sub


Public Sub setRowsHeight()
On Error GoTo Err_setRowsHeight

  Application.Cursor = xlWait
  Application.ScreenUpdating = False
  
  shMain.Rows(conf_Props.RSTART & ":" & conf_Props.REND).ROWHEIGHT = conf_Props.ROWHEIGHT

Exit_setRowsHeight:
  Application.Cursor = xlDefault
  Application.ScreenUpdating = True
  Exit Sub
  
Err_setRowsHeight:
  myLogger.LogError Err.Number & ": " & Err.Source & " / " & Err.Description, "Utils.setRowHeight"
  Resume Exit_setRowsHeight
End Sub

Public Sub setZoom(ByVal zoom As Long)
On Error GoTo Err_setZoom
    
  Application.Cursor = xlWait
  Application.ScreenUpdating = False
  
  ActiveWindow.zoom = zoom
    
Exit_setZoom:
  Application.Cursor = xlDefault
  Application.ScreenUpdating = True
  Exit Sub
  
Err_setZoom:
  myLogger.LogError Err.Number & ": " & Err.Source & " / " & Err.Description, "Utils.setZoom"
  Resume Exit_setZoom
End Sub

Public Sub setFreezePanel(ByVal Sh As Worksheet, ByVal cell As Variant)
On Error GoTo Err_setFreezePanel

  Application.Cursor = xlWait
  Application.ScreenUpdating = False
  
  ActiveWindow.FreezePanes = False
  
  Sh.Range(cell).Select
  ActiveWindow.FreezePanes = True

Exit_setFreezePanel:
  Application.Cursor = xlDefault
  Application.ScreenUpdating = True
  Exit Sub

Err_setFreezePanel:
  myLogger.LogError Err.Number & ": " & Err.Source & " / " & Err.Description, "Utils.setFreezePanel"
  Resume Exit_setFreezePanel
End Sub

Public Sub setProtection(ByVal Sh As Worksheet, ByVal onoff As Boolean)
On Error GoTo Err_setProtection

  Application.Cursor = xlWait
  Application.ScreenUpdating = False
  
  If onoff Then
    Sh.Protect Password:=conf_Props.SHEETPASSWORD, DrawingObjects:=True, _
                   Contents:=True, Scenarios:=True, UserInterfaceOnly:=True, _
                   AllowFormattingCells:=False, AllowFormattingColumns:=True, _
                   AllowFormattingRows:=True, AllowInsertingColumns:=False, AllowInsertingRows:=False, _
                   AllowInsertingHyperlinks:=False, AllowDeletingColumns:=False, AllowDeletingRows:=False, _
                   AllowSorting:=True, AllowFiltering:=True, AllowUsingPivotTables:=True
  Else
    Sh.Unprotect Password:=conf_Props.SHEETPASSWORD
  End If

Exit_setProtection:
  Application.Cursor = xlDefault
  Application.ScreenUpdating = True
  Exit Sub

Err_setProtection:
  myLogger.LogError Err.Number & ": " & Err.Source & " / " & Err.Description, "Utils.setProtection"
  Resume Exit_setProtection
End Sub

Public Function startFromSP() As Boolean
  startFromSP = Left$(ActiveWorkbook.Path, 4) = "http"
End Function
