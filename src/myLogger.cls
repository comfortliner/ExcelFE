VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "myLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("Development")
'@PredeclaredId
Option Explicit

Private Type TLogger
  LOGENABLED As Boolean
  LOGTHRESHOLD As LogLevel
End Type

Private This As TLogger

Private Sub Class_Initialize()
  This.LOGENABLED = conf_Props.LOGENABLED
  This.LOGTHRESHOLD = conf_Props.LOGTHRESHOLD
End Sub

'@Ignore ProcedureNotUsed
Public Property Get Enabled() As Boolean
  Enabled = This.LOGENABLED
End Property

'@Ignore ProcedureNotUsed
Public Property Let Enabled(ByVal value As Boolean)
  This.LOGENABLED = value
End Property

'@Ignore ProcedureNotUsed
Public Property Get Threshold() As LogLevel
  Threshold = This.LOGTHRESHOLD
End Property

'@Ignore ProcedureNotUsed
Public Property Let Threshold(ByVal value As LogLevel)
  This.LOGTHRESHOLD = value
End Property

'@Ignore ProcedureNotUsed
Public Sub LogTrace(ByVal Message As String, Optional ByVal From As String = vbNullString)
  Me.Log bTrace, Message, From
End Sub

'@Ignore ProcedureNotUsed
Public Sub LogDebug(ByVal Message As String, Optional ByVal From As String = vbNullString)
  Me.Log cDebug, Message, From
End Sub

'@Ignore ProcedureNotUsed
Public Sub LogInfo(ByVal Message As String, Optional ByVal From As String = vbNullString)
  Me.Log dInfo, Message, From
End Sub

'@Ignore ProcedureNotUsed
Public Sub LogWarn(ByVal Message As String, Optional ByVal From As String = vbNullString)
  Me.Log eWarn, Message, From
End Sub

'@Ignore ProcedureNotUsed
Public Sub LogError(ByVal Message As String, Optional ByVal From As String = vbNullString)
  Me.Log fError, Message, From
End Sub

'@Ignore ProcedureNotUsed
Public Sub Log(ByVal Level As LogLevel, ByVal Message As String, ByVal From As String)
  Dim log_LevelValue As String
  Dim log_LevelColor As RenderingColor
  Dim rng As Range

  If This.LOGTHRESHOLD = aOff Or Level < This.LOGTHRESHOLD Then
    Exit Sub
  End If

  Select Case Level
  Case bTrace
    log_LevelValue = "Trace"
    log_LevelColor = White
  Case cDebug
    log_LevelValue = "Debug"
    log_LevelColor = White
  Case dInfo
    log_LevelValue = "Info"
    log_LevelColor = Green
  Case eWarn
    log_LevelValue = "WARN"
    log_LevelColor = Yellow
  Case fError
    log_LevelValue = "ERROR"
    log_LevelColor = Red
  End Select

  Set rng = shMain.Range(conf_Props.LOGOUTPUT)
  rng.FormulaR1C1 = log_LevelValue
  rng.Interior.Color = log_LevelColor
  
  Set rng = shMain.Range(conf_Props.LOGOUTPUT).Offset(1, 0)
  rng.FormulaR1C1 = Format$(CDate(Now), "hh:mm:ss") & _
                    " - " & IIf(From <> vbNullString, From & _
                    ": ", vbNullString) & Message

  If Level < dInfo Then
    Debug.Print log_LevelValue & " " & _
                Format$(CDate(Now), "hh:mm:ss") & _
                " - " & IIf(From <> vbNullString, From & _
                ": ", vbNullString) & Message
  End If
End Sub

