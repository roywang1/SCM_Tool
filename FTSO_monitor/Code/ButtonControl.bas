Attribute VB_Name = "ButtonControl"
Private Sub btnPTI_Click()
    ' 呼叫處理 PTI 工作表的子程�
    CreateDictionariesForAllSheets
    ProcessSheet "PTI"
End Sub

Private Sub btnASE_Click()
    ' 呼叫處理 ASE 工作表的子程序
    CreateDictionariesForAllSheets
    ProcessSheet "ASE"
End Sub

Private Sub btnSigurd_Click()
    ' 呼叫處理 Sigurd 工作表的子程�
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
    
    
    ' 監控 I3 和 AN1 的變化
    If Not Intersect(Target, ws.Range(MonthValue)) Is Nothing Then
        
        ' 檢� dAN5 的值並觸發相應的按鈕
        ' 如果 I3 的值改變
        If ws.Range(MonthValue).Value <> lastMonthValue Then
            Select Case ws.Range(HiddenRefValue).Value
                Case "PTI"
                    Debug.Print "###btnPTI trigger"
                    Call btnPTI_Click ' 假設 Button1 是你的按鈕名稱
                Case "ASE"
                    Debug.Print "###btnASE trigger"
                    Call btnASE_Click ' 假設 Button1 是你的按鈕名稱
                Case "Sigurd"
                    Debug.Print "###btnSigurd trigger"
                    Call btnSigurd_Click ' 假設 Button1 是你的按鈕名稱
            End Select
        End If

        ' 更新 lastI3Value 和 lastAN1Value
        lastMonthValue = ws.Range(MonthValue).Value
        lastHiddenRefValue = ws.Range(HiddenRefValue).Value

    End If
End Sub

