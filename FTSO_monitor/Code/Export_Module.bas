Attribute VB_Name = "Export_Module"
' Excel macro to export all VBA source code in this project to text files for proper source control versioning
' Requires enabling the Excel setting in Options/Trust Center/Trust Center Settings/Macro Settings/Trust access to the VBA project object model
Public Sub ExportVisualBasicCode()
    Const Module = 1
    Const ClassModule = 2
    Const Form = 3
    Const Document = 100
    Const Padding = 24
    
    Dim VBComponent As Object
    Dim count As Integer
    Dim path As String
    Dim directory As String
    Dim extension As String
    'Dim fso As New FileSystemObject
    
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    directory = Application.Run("Version_Check.Get_VBA_Repo")
    If Not fso.FolderExists(directory) Then
        Call fso.CreateFolder(directory)
    End If
    
    directory = Application.Run("Version_Check.Get_VBA_Path")
    
    If Not fso.FolderExists(directory) Then
        Call fso.CreateFolder(directory)
    End If
    
    ''''''''''''''''''''''''''''''''''
    'Clean the folder with Macro first
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oFolder = oFSO.GetFolder(directory)
    
    On Error Resume Next
    For Each oFile In oFolder.Files
        file_path = oFolder & "\" & oFile.Name
        Kill file_path
    Next
    On Error GoTo 0
    '''''''''''''''''''''''''''''''''''

    count = 0
    
    Set fso = Nothing
    
    For Each VBComponent In ThisWorkbook.VBProject.VBComponents
        Select Case VBComponent.Type
            Case ClassModule, Document
                extension = ".cls"
            Case Form
                extension = ".frm"
            Case Module
                extension = ".bas"
            Case Else
                extension = ".txt"
        End Select
            
                
        On Error Resume Next
        Err.Clear
        
        path = directory & "\" & VBComponent.Name & extension
        If extension = ".bas" Then
            Call VBComponent.Export(path)
            
            If Err.Number <> 0 Then
                Call MsgBox("Failed to export " & VBComponent.Name & " to " & path, vbCritical)
            Else
                count = count + 1
                Debug.Print "Exported " & Left$(VBComponent.Name & ":" & Space(Padding), Padding) & path
            End If
            On Error GoTo 0
        End If

    Next
    
    ThisWorkbook.Save
'    Application.StatusBar = "Successfully exported " & CStr(count) & " VBA files to " & directory
'    Application.OnTime Now + TimeSerial(0, 0, 10), "ClearStatusBar"
End Sub
