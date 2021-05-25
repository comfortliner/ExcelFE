VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "myDBDefaultParam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Model.DB")
Option Explicit

Implements ImyDBDefaultParam

Private Type TDBParam
  PName As String
  PType As DataTypeEnum
  PDirection As ParameterDirectionEnum
  PSize As Long
  pValue As Variant
End Type

Private This As TDBParam

Private Property Get ImyDBDefaultParam_pName() As String
  ImyDBDefaultParam_pName = This.PName
End Property

Private Property Get ImyDBDefaultParam_pType() As DataTypeEnum
  ImyDBDefaultParam_pType = This.PType
End Property

Private Property Get ImyDBDefaultParam_pDirection() As ParameterDirectionEnum
  ImyDBDefaultParam_pDirection = This.PDirection
End Property

Private Property Get ImyDBDefaultParam_pSize() As Long
  ImyDBDefaultParam_pSize = This.PSize
End Property

Private Property Get ImyDBDefaultParam_pValue() As Variant
  ImyDBDefaultParam_pValue = This.pValue
End Property

Public Sub FillData( _
       ByVal PName As String, _
       ByVal PType As DataTypeEnum, _
       ByVal PDirection As ParameterDirectionEnum, _
       Optional ByVal PSize As Long, _
       Optional ByVal pValue As Variant _
)
    
  This.PName = PName
  This.PType = PType
  This.PDirection = PDirection
  This.PSize = PSize
  This.pValue = pValue
End Sub

