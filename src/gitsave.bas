Attribute VB_Name = "gitsave"
'@Folder("Development")
Option Explicit
' Library Reference used: Microsoft Visual Basic for Applications Extensibility 5.3 / vbe6ext.olb

Public Sub run()
    
  DeleteAndMake
  ExportModules

End Sub

Public Sub DeleteAndMake()
        
  Dim fso As Object
  Set fso = CreateObject("Scripting.FileSystemObject")

  Dim parentFolder1 As String: parentFolder1 = ThisWorkbook.Path & "\src_github"
  Dim parentFolder2 As String: parentFolder2 = ThisWorkbook.Path & "\src_" & Left$(ThisWorkbook.name, (InStrRev(ThisWorkbook.name, ".", -1, vbTextCompare) - 1))
        
  On Error Resume Next
  fso.DeleteFolder parentFolder1
  fso.DeleteFolder parentFolder2
  On Error GoTo 0
    
  MkDir parentFolder1
  MkDir parentFolder2
    
End Sub

Public Sub ExportModules()
       
  Dim pathToExport1 As String: pathToExport1 = ThisWorkbook.Path & "\src_github"
  Dim pathToExport2 As String: pathToExport2 = ThisWorkbook.Path & "\src_" & Left$(ThisWorkbook.name, (InStrRev(ThisWorkbook.name, ".", -1, vbTextCompare) - 1))
    
  If Dir(pathToExport1) <> vbNullString Then
    Kill pathToExport1 & "*.*"
  End If

  If Dir(pathToExport2) <> vbNullString Then
    Kill pathToExport2 & "*.*"
  End If
  
  Dim wkb As Workbook: Set wkb = Excel.Workbooks.Item(ThisWorkbook.name)
    
  Dim file As String
  Dim component As VBIDE.VBComponent
  Dim tryExport As Long

  For Each component In wkb.VBProject.VBComponents
    tryExport = 1
    file = component.name
       
    Select Case component.Type
    Case vbext_ct_ClassModule
      file = file & ".cls"
    Case vbext_ct_MSForm
      file = file & ".frm"
    Case vbext_ct_StdModule
      file = file & ".bas"
    Case vbext_ct_Document
      file = file & ".doccls"
    Case Else
      tryExport = 0
    End Select
    
    
    ' Not to check into version control system
    If InStr(file, "conf") = 1 Then tryExport = 2
    If InStr(file, "cust") = 1 Then tryExport = 2
    If InStr(file, "Tabelle") = 1 Then tryExport = 2
    If InStr(file, "DieseArbeitsmappe") = 1 Then tryExport = 2
    
    If tryExport = 1 Then
      Debug.Print "Exporting " & file
      component.Export pathToExport1 & "\" & file
    End If
    
    If tryExport = 2 Then
      Debug.Print "Exporting " & file
      component.Export pathToExport2 & "\" & file
    End If
  Next

  Debug.Print "Exported at " & pathToExport1 & " and " & pathToExport2
    
End Sub

