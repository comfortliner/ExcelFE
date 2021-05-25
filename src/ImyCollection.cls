VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ImyCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Model.Collection")
Option Explicit

Public Property Get columnsCollection() As VBA.Collection
End Property

Public Property Get paramsCollection() As VBA.Collection
End Property

Public Property Get formcondCollection() As VBA.Collection
End Property

Public Sub addColumn(ByVal column As ImyColumn)
End Sub

Public Sub addSlicer(ByVal Slicer As ImySlicer)
End Sub

Public Sub addButton(ByVal Button As ImyButton)
End Sub

Public Sub adddbDefaultParam(ByVal Param As ImyDBDefaultParam)
End Sub

Public Sub addFormCond(ByVal FormCond As ImyFormCond)
End Sub

Public Sub generateshMain()
End Sub

Public Sub initializeshMain()
End Sub

