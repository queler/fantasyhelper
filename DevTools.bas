Attribute VB_Name = "DevTools"
Option Explicit
Private Declare Function SHGetSpecialFolderPathW Lib "shell32.dll" ( _
    ByVal hwndOwner As Long, _
    ByVal lpszPath As Long, _
    ByVal nFolder As Long, _
    ByVal fCreate As Long) As Boolean

Private Const MD As Long = &H5
 Public Function My_Documents() As String


    Dim Folders(512) As Byte
    
    My_Documents = vbNullString
    
    If SHGetSpecialFolderPathW(0&, _
        ByVal VarPtr(Folders(0)), _
            MD, 0&) Then
        My_Documents = Left$(Folders, InStr(Folders, Chr$(0)) - 1)
    End If
End Function
Sub exp()
ExportSourceFiles My_Documents & "\fh\"
End Sub
Public Sub ImportSourceFiles(sourcePath As String)
    Dim file As String
    file = Dir(sourcePath)
    While (file <> "")
        Application.VBE.ActiveVBProject.VBComponents.Import sourcePath & file
        file = Dir
    Wend
End Sub

Public Sub ExportSourceFiles(destPath As String)
     
    Dim component As VBComponent
    For Each component In Application.VBE.ActiveVBProject.VBComponents
        If component.Type = vbext_ct_ClassModule Or component.Type = vbext_ct_StdModule Then
            component.Export destPath & component.name & ToFileExtension(component.Type)
        End If
    Next
     
End Sub

Public Sub RemoveAllModules()
    Dim project As VBProject
    Set project = Application.VBE.ActiveVBProject
    
    Dim comp As VBComponent
    For Each comp In project.VBComponents
        If Not comp.name = "DevTools" And (comp.Type = vbext_ct_ClassModule Or comp.Type = vbext_ct_StdModule) Then
            project.VBComponents.Remove comp
        End If
    Next
End Sub

Private Function ToFileExtension(vbeComponentType As vbext_ComponentType) As String

    Select Case vbeComponentType
        Case vbext_ComponentType.vbext_ct_ClassModule
            ToFileExtension = ".cls"
        Case vbext_ComponentType.vbext_ct_StdModule
            ToFileExtension = ".bas"
        Case vbext_ComponentType.vbext_ct_MSForm
            ToFileExtension = ".frm"
        Case vbext_ComponentType.vbext_ct_ActiveXDesigner
        Case vbext_ComponentType.vbext_ct_Document
        Case Else
            ToFileExtension = vbNullString
    End Select
    
End Function

