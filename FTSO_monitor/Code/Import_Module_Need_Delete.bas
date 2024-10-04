Attribute VB_Name = "Import_Module_Need_Delete"
Sub ImportVisualBasicCode()
 
    Dim oFSO As Object
    Dim oFolder As Object
    Dim oFile As Object
    Dim i As Integer
    Dim directory As String
                 
         
    Set oFSO = CreateObject("Scripting.FileSystemObject")
     
    'Set oFolder = oFSO.GetFolder(ActiveWorkbook.path & "\VisualBasic")
    VBA_DIR = Application.Run("Version_Check.Get_VBA_Path")
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(VBA_DIR) Then
        Debug.Print "Reference Macro code path1 not exists..."
        
        check_path = False
                
        If check_path = False Then
            Exit Sub
        End If
        
    End If
    
    
    Set oFolder = oFSO.GetFolder(VBA_DIR)
    
    '''''''''''''
    'Check Macro version first, then decide if needs to do replacement
    Cur_Version = Application.Run("Version_Check.Get_Version")
    
    Debug.Print Cur_Version
    
    Version_File = oFolder & "\" & "Version_Check.bas"
    strFileExists = Dir(Version_File)
    If strFileExists = "" Then
        Debug.Print "Reference Macro VBA (Version_Check) not exists..."
        Exit Sub
    End If
    
    
    ' 先備份原來的模組
    On Error Resume Next ' 忽略錯誤，因為模組可能不存在
    Set originalModule = ThisWorkbook.VBProject.VBComponents("Version_Check")
    On Error GoTo 0 ' 恢復錯誤處理
    
    If Not originalModule Is Nothing Then
        ' 導出原始模組到臨時檔案
        Filename = "Version_Check_" & Format(Now, "yyyymmdd_hhnnss") & ".bas"
        tempFilePath = ThisWorkbook.path & "\" & Filename
        originalModule.Export tempFilePath
        Debug.Print "Backup original module to: " & tempFilePath
    End If
    
    
    ' 刪除當前版本檢查模組
    If Not originalModule Is Nothing Then
        ThisWorkbook.VBProject.VBComponents.Remove originalModule
        Debug.Print "Remove old module: Version_Check"
    End If
    
    ' 導入新的版本檢查模組
    ThisWorkbook.VBProject.VBComponents.Import Version_File
    Debug.Print "Import module: Version_Check"
    
    ' 讀取新導入的版本
    Import_Version = Application.Run("Version_Check.Get_Version")
    Debug.Print "Imported Version: " & Import_Version
    
    ' 比較當前版本與導入版本
    If Cur_Version <> Import_Version Then
        ' 版本不相同，詢問用戶是否要更新
        Dim userResponse As VbMsgBoxResult
        userResponse = MsgBox("The versions are different. Do you want to update the version?", vbQuestion + vbYesNo, "Confirm Update")
        
        ' 根據用戶的選擇執行相應的操作
        If userResponse = vbYes Then
            Macro_Update = True
            Debug.Print "Updating Macro code..."
        Else
            Macro_Update = False

            ' 刪除剛導入的模組
            ThisWorkbook.VBProject.VBComponents.Remove ThisWorkbook.VBProject.VBComponents("Version_Check")
            Debug.Print "Removed the newly imported module."
            
            ' 如果用戶選擇不更新，則恢復原來的模組
            If Not originalModule Is Nothing Then
                ThisWorkbook.VBProject.VBComponents.Import tempFilePath
                Debug.Print "User chose not to update. Restored the original module."
            End If
            
        End If
    Else
        Macro_Update = False
        Debug.Print "The Macro code is the latest!"
    End If

    ' ??????
    On Error Resume Next ' ????,???????
    Kill tempFilePath
    On Error GoTo 0 ' ??????
    Debug.Print "Temporary file deleted: " & tempFilePath
    '''''''''''''
    
    
    If Macro_Update = True Then
    
        Set Module_List = CreateObject("System.Collections.ArrayList")
        
        For i = 1 To ThisWorkbook.VBProject.VBComponents.count
            'if not ThisWorkbook.VBProject.VBComponents(i).type = ".bas"
            If ThisWorkbook.VBProject.VBComponents(i).Type = 1 Then
                
                Module_Name = ThisWorkbook.VBProject.VBComponents(i).CodeModule.Name
                Module_List.Add Module_Name
            End If
            
        Next
        
        
        For Each j In Module_List
            If Not j = "Import_Module" Then
                ThisWorkbook.VBProject.VBComponents.Remove ThisWorkbook.VBProject.VBComponents(j)
                Debug.Print "Remove old module: " & j
            Else
                'rename "Import_Module" first, then delete it in the last
                ThisWorkbook.VBProject.VBComponents(j).Name = j & "_Need_Delete"
            End If
        Next
        
        
        For Each oFile In oFolder.Files
         
            directory = oFolder & "\" & oFile.Name
            Filename_without_ext = oFSO.GetBaseName(oFile)
        
            On Error Resume Next
            
    
            ThisWorkbook.VBProject.VBComponents.Import directory
            Debug.Print "Import module: " & Filename_without_ext
            
            
    '        If Err.Number <> 0 Then
    '            Call MsgBox("Failed to import " & oFile.Name, vbCritical)
    '        End If
        Next oFile
    End If
End Sub
