Attribute VB_Name = "ButtonControl"
Private Sub btnPTI_Click()
    ' �I�s�B�z PTI �u�@���l�{�
    CreateDictionariesForAllSheets
    ProcessSheet "PTI"
End Sub

Private Sub btnASE_Click()
    ' �I�s�B�z ASE �u�@���l�{��
    CreateDictionariesForAllSheets
    ProcessSheet "ASE"
End Sub

Private Sub btnSigurd_Click()
    ' �I�s�B�z Sigurd �u�@���l�{�
    CreateDictionariesForAllSheets
    ProcessSheet "Sigurd"
End Sub



Private Sub Worksheet_Change(ByVal Target As Range)
    Dim lastMonthValue As Variant
    Dim lastHiddenRefValue As Variant
    Dim MonthValue As Variant
    Dim HiddenRefValue As Variant
    
    MonthValue = "F3"
    HiddenRefValue = "AK5"
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("SO Summary")
    
    
    ' �ʱ� I3 �M AN1 ���ܤ�
    If Not Intersect(Target, ws.Range(MonthValue)) Is Nothing Then
        
        ' �ˬ dAN5 ���Ȩ�Ĳ�o���������s
        ' �p�G I3 ���ȧ���
        If ws.Range(MonthValue).Value <> lastMonthValue Then
            Select Case ws.Range(HiddenRefValue).Value
                Case "PTI"
                    Debug.Print "###btnPTI trigger"
                    Call btnPTI_Click ' ���] Button1 �O�A�����s�W��
                Case "ASE"
                    Debug.Print "###btnASE trigger"
                    Call btnASE_Click ' ���] Button1 �O�A�����s�W��
                Case "Sigurd"
                    Debug.Print "###btnSigurd trigger"
                    Call btnSigurd_Click ' ���] Button1 �O�A�����s�W��
            End Select
        End If

        ' ��s lastI3Value �M lastAN1Value
        lastMonthValue = ws.Range(MonthValue).Value
        lastHiddenRefValue = ws.Range(HiddenRefValue).Value

    End If
End Sub

