Attribute VB_Name = "SlicerFactory"
'@Folder("Controller")
Option Explicit

Public Function Create( _
       ByVal StartTileX As Long, _
       ByVal StartTileY As Long, _
       ByVal StepTilesX As Long, _
       ByVal StepTilesY As Long, _
       ByVal Source As String, _
       ByVal SourceColumn As String, _
       ByVal SliderHeader As String, _
       Optional ByVal NumOfCols As Long = 1 _
       ) As ImySlicer
  
  Dim NewSlicer As mySlicer
  Set NewSlicer = New mySlicer
  
  NewSlicer.FillData StartTileX, StartTileY, StepTilesX, StepTilesY, Source, _
                     SourceColumn, SliderHeader, NumOfCols
  Set Create = NewSlicer
End Function

Public Sub clearFilter()
  Dim list As ListObject

  For Each list In ActiveSheet.ListObjects
    If list.AutoFilter.FilterMode Then
      list.AutoFilter.ShowAllData
    End If
  Next
End Sub

Public Sub deleteAllSlicer()
  Dim slc As SlicerCache
  
  For Each slc In ActiveWorkbook.SlicerCaches
    slc.Delete
  Next slc
  
End Sub


