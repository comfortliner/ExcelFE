VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "myMSSQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Controller")
Option Explicit
' Library Reference used: Microsoft ActiveX Data Objects 6.1 Library / msado15.dll

Private Type db
  DBProvider As String
  DBDataSource As String
  DBInitialCatalog As String
  DBUserID As String
  DBPassword As String
  DBSPName As String
End Type

Private This As db

Private Sub Class_Initialize()

  This.DBProvider = conf_Props.DBProvider
  This.DBDataSource = conf_Props.DBDataSource
  This.DBInitialCatalog = conf_Props.DBInitialCatalog
  This.DBUserID = conf_Props.DBUserID
  This.DBPassword = conf_Props.DBPassword
  This.DBSPName = conf_Props.DBSPName
    
End Sub

Public Sub execute(ByVal row As Long)
On Error GoTo Err_execute

  Dim cnn As ADODB.Connection
  Set cnn = New ADODB.Connection
  
  Dim cmd As ADODB.Command
  Set cmd = New ADODB.Command
  
  Dim rs As ADODB.Recordset
  Set rs = New ADODB.Recordset
  
  Dim cnnStr As String

  Dim column As ImyColumn
  Set column = New myColumn
  
  Dim parameter As ImyDBDefaultParam
  Set parameter = New myDBDefaultParam
  

  myLogger.LogWarn "Werte werden in die DB geschrieben..", "MSSQL.execute"
  
  Application.Cursor = xlWait
  Application.ScreenUpdating = False

  cnnStr = "Provider=" & This.DBProvider & ";" & _
           "Data Source=" & This.DBDataSource & ";" & _
           "Initial Catalog=" & This.DBInitialCatalog & ";" & _
           "User ID=" & This.DBUserID & ";" & _
           "Password=" & This.DBPassword & ";"
  
  With cnn
    .CommandTimeout = 900
    .ConnectionString = cnnStr
    .Open
  End With
  
  With cmd
    '@Ignore ImplicitDefaultMemberAccess
    .ActiveConnection = cnn
    .CommandType = adCmdStoredProc
    .CommandText = This.DBSPName
    .NamedParameters = True
    
    For Each parameter In Collection.paramsCollection
      .Parameters.Append .createParameter(parameter.PName, parameter.PType, parameter.PDirection, parameter.PSize, parameter.pValue)
    Next parameter
        
    For Each column In Collection.columnsCollection
      If column.PName <> vbNullString Then _
        .Parameters.Append .createParameter(column.PName, column.PType, column.PDirection, column.PSize, shMain.Cells(row, column.columnChar).value)
    Next column
  
  End With
  
  If conf_Props.LOGTHRESHOLD < dInfo Then
    Dim i As Long
    Dim str As String
    Debug.Print cmd.CommandText
    For i = 0 To cmd.Parameters.Count - 1
      str = str & vbTab & i & ":" & cmd.Parameters.Item(i).value
    Next i
    Debug.Print str
  Else
    With rs
      .CursorType = adOpenStatic
      .CursorLocation = adUseClient
      .LockType = adLockOptimistic
      .Open cmd
    End With
  
    Set rs = cmd.execute
  End If

  Application.Cursor = xlDefault
  Application.ScreenUpdating = True
  
  myLogger.LogInfo "Werte der Zeile " & row & " wurden in der Datenbank gespeichert...", "MSSQL.execute"


Exit_execute:
  cnn.Close
  Set cmd = Nothing
  Set rs = Nothing
  Application.Cursor = xlDefault
  Application.ScreenUpdating = True
  
  Exit Sub
  
Err_execute:
  myLogger.LogError Err.Number & ": " & Err.Source & " / " & Err.Description, "MSSQL.execute"
  Resume Exit_execute
End Sub
