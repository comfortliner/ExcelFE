Attribute VB_Name = "DBParamFactory"
'@Folder("Controller")
Option Explicit

Public Function Create( _
       ByVal PName As String, _
       ByVal PType As DataTypeEnum, _
       ByVal PDirection As ParameterDirectionEnum, _
       Optional ByVal PSize As Long, _
       Optional ByVal pValue As Variant _
       ) As ImyDBDefaultParam
  
  Dim NewParam As myDBDefaultParam
  Set NewParam = New myDBDefaultParam
  
  NewParam.FillData PName, PType, PDirection, PSize, pValue
  Set Create = NewParam
End Function

