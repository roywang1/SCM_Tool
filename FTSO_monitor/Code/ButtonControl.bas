Attribute VB_Name = "ButtonControl"
Private Sub btnPTI_Click()
    ' ©I¥s³B²z PTI ¤u§@ªíªº¤lµ{§
    CreateDictionariesForAllSheets
    ProcessSheet "PTI"
End Sub

Private Sub btnASE_Click()
    ' ©I¥s³B²z ASE ¤u§@ªíªº¤lµ{§Ç
    CreateDictionariesForAllSheets
    ProcessSheet "ASE"
End Sub

Private Sub btnSigurd_Click()
    ' ©I¥s³B²z Sigurd ¤u§@ªíªº¤lµ{§
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
    
    
    ' ºÊ±± I3 ©M AN1 ªºÅÜ¤Æ
    If Not Intersect(Target, ws.Range(MonthValue)) Is Nothing Then
        
        ' ÀË¬ dAN5 ªº­È¨ÃÄ²µo¬ÛÀ³ªº«ö¶s
        ' ¦pªG I3 ªº­È§ïÅÜ
        If ws.Range(MonthValue).Value <> lastMonthValue Then
            Select Case ws.Range(HiddenRefValue).Value
                Case "PTI"
                    Debug.Print "###btnPTI trigger"
                    Call btnPTI_Click ' °²³] Button1 ¬O§Aªº«ö¶s¦WºÙ
                Case "ASE"
                    Debug.Print "###btnASE trigger"
                    Call btnASE_Click ' °²³] Button1 ¬O§Aªº«ö¶s¦WºÙ
                Case "Sigurd"
                    Debug.Print "###btnSigurd trigger"
                    Call btnSigurd_Click ' °²³] Button1 ¬O§Aªº«ö¶s¦WºÙ
            End Select
        End If

        ' §ó·s lastI3Value ©M lastAN1Value
        lastMonthValue = ws.Range(MonthValue).Value
        lastHiddenRefValue = ws.Range(HiddenRefValue).Value

    End If
End Sub

