Attribute VB_Name = "Version_Check"
Private Function Get_Version() As String
    
    Get_Version = "0.1/1004 2024"
    
    '10/04: Initial Version
        
End Function

Private Function Get_VBA_Repo() As String

    Get_VBA_Repo = ThisWorkbook.path & "\Code"

End Function

Private Function Get_VBA_Path() As String
    
    Get_VBA_Path = ThisWorkbook.path & "\Code"

End Function
