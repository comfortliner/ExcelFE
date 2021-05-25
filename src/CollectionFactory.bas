Attribute VB_Name = "CollectionFactory"
'@Folder("Controller")
Option Explicit

Public Function Create( _
       ) As ImyCollection
  
  Dim NewCollection As myCollection
  Set NewCollection = New myCollection

  'NewCollection.FillData

  Set Create = NewCollection
End Function


