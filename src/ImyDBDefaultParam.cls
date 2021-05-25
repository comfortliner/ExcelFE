VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ImyDBDefaultParam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Model.DB")
Option Explicit

Public Property Get PName() As String
End Property

Public Property Get PType() As DataTypeEnum
End Property

Public Property Get PDirection() As ParameterDirectionEnum
End Property

Public Property Get PSize() As Long
End Property

Public Property Get pValue() As Variant
End Property

