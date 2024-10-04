Attribute VB_Name = "Main"
Dim dictPTI As Object
Dim dictASE As Object
Dim dictSigurd As Object


Sub ProcessSheet(sheetName As String)
    Dim dict As Object
    
    ' �ھڤu�@��W�ٿ�ܹ������r��
    Select Case sheetName
        Case "PTI"
            If dictPTI Is Nothing Then
                MsgBox "PTI dictionary is not initialized!"
                Exit Sub
            End If
            Set dict = dictPTI
        Case "ASE"
            If dictASE Is Nothing Then
                MsgBox "ASE dictionary is not initialized!"
                Exit Sub
            End If
            Set dict = dictASE
        Case "Sigurd"
            If dictSigurd Is Nothing Then
                MsgBox "Sigurd dictionary is not initialized!"
                Exit Sub
            End If
            Set dict = dictSigurd
        Case Else
            MsgBox "Invalid sheet name: " & sheetName
            Exit Sub
    End Select
    
    ' ���L�r���������u�@��
    PrintDictionaryToSheet dict, sheetName, "SO Summary"
End Sub



Sub CreateDictionariesForAllSheets()
    ' �Ыبæs�x PTI, ASE, Sigurd �T�Ӥ������r��
    Set dictPTI = CreateDictionaryForSheet("PTI")
    Set dictASE = CreateDictionaryForSheet("ASE")
    Set dictSigurd = CreateDictionaryForSheet("Sigurd")
    
    ' �ոտ�X�A�T�{�r��w�Ы�
    Debug.Print "PTI Dictionary Keys: " & dictPTI.count
    Debug.Print "ASE Dictionary Keys: " & dictASE.count
    Debug.Print "Sigurd Dictionary Keys: " & dictSigurd.count
    
    
    ' PrintDictionaryToSheet dictPTI, "PTI", "SO Summary"
    ' PrintDictionaryToSheet dictASE, "ASE", "SO Summary"
    ' PrintDictionaryToSheet dictSigurd, "Sigurd", "SO Summary"
    
    
    ' MsgBox "Dictionaries created for PTI, ASE, and Sigurd!"
End Sub


Function CreateDictionaryForSheet(sheetName As String) As Object
    Dim ws As Worksheet
    Dim dict As Object
    Dim subDict As Object
    Dim lastRow As Long
    Dim rowNum As Long
    Dim key As String
    Dim dateKey As Variant
    Dim bValue As Variant
    Dim cValue As Variant
    Dim eValue As Variant
    Dim fValue As Variant
    Dim gValue As Variant
    Dim hValue As Variant
    Dim iValue As Variant
    Dim dateList As Collection
    Dim sortedDates As Variant
    Dim i As Long
    
    ' �]�m�u�@��
    Set ws = ThisWorkbook.Sheets(sheetName)
    
    ' �ЫإD�r��
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' ���̫�@��
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    
    ' �Ыؤ�����X
    Set dateList = New Collection
    
    ' �M���C�@��A�q�ĤG��}�l
    For rowNum = 2 To lastRow
        dateKey = ws.Cells(rowNum, 1).Value ' A �檺�ȧ@����� key
        
        ' �N����[�J���X���]�קK���ơ^
        On Error Resume Next
        dateList.Add dateKey, CStr(dateKey)
        On Error GoTo 0
    Next rowNum
    
    ' �N������X�ഫ���ƲըñƧ�
    sortedDates = SortCollection(dateList)

    ' �M���Ƨǫ᪺���
    For i = LBound(sortedDates) To UBound(sortedDates)
        dateKey = sortedDates(i)
        
        ' �A���M���C�@��H��R�r��
        For rowNum = 2 To lastRow
            key = ws.Cells(rowNum, 4).Value ' D �檺�ȧ@���D�r�媺 key
            
            ' �u�B�z�ŦX��e�������
            If ws.Cells(rowNum, 1).Value = dateKey Then
                bValue = ws.Cells(rowNum, 2).Value ' B �檺��
                cValue = ws.Cells(rowNum, 3).Value ' C �檺��
                eValue = ws.Cells(rowNum, 5).Value ' E �檺��
                fValue = ws.Cells(rowNum, 6).Value ' F �檺��
                
                gValue = ws.Cells(rowNum, 7).Value ' G �檺��
                hValue = ws.Cells(rowNum, 8).Value ' H �檺��
                iValue = ws.Cells(rowNum, 9).Value ' I �檺��
                
                ' �p�G�D�r�夤���s�b�� key�A�h�Ыطs���l�r��
                If Not dict.Exists(key) Then
                    Set subDict = CreateObject("Scripting.Dictionary")
                    dict.Add key, subDict
                Else
                    Set subDict = dict(key) ' �p�G key �w�g�s�b�A���X�l�r��
                End If
                
                ' �ˬd�l�r�夤�O�_�w������@�� key
                If Not subDict.Exists(dateKey) Then
                    ' �N B, C, E, F �檺�Ȧs�J�ƾڦC��
                    subDict.Add dateKey, Array(bValue, cValue, eValue, fValue, gValue, hValue, iValue)
                Else
                    ' �i��G�B�z�ۦP��������p�A�o�̤��B�z���Ƥ��
                End If
            End If
        Next rowNum
    Next i
    
    ' ��^�Ыت��r��
    Set CreateDictionaryForSheet = dict
End Function

Function SortCollection(col As Collection) As Variant
    Dim arr() As Variant
    Dim i As Long, j As Long
    Dim temp As Variant

    ' �N���X�ഫ���Ʋ�
    ReDim arr(1 To col.count)
    For i = 1 To col.count
        arr(i) = col(i)
    Next i

    ' �_�w�Ƨ�
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(i) > arr(j) Then
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            End If
        Next j
    Next i

    SortCollection = arr
End Function



Function PrintDictionaryToSheet(ByVal dict As Object, OSAT As String, sheetName As String)
    Dim ws As Worksheet
    Dim subDict As Object
    Dim key As Variant
    Dim dateKey As Variant
    Dim rowNum As Long
    Dim monthFilter As Integer
    Dim dayCol As Integer
    Dim dateParts() As String
    Dim targetMonth As Integer
    Dim targetDay As Integer
    Dim prevdateKey As Variant
    Dim no_month_limit_prevdateKey As Variant
    
    Dim dayCol_start As Integer
    Dim hidden_ref_value_col As Integer
    Dim MonthValue As Variant
    Dim Item_col As Integer
    Dim FAB_col As Integer
    Dim Nick_name_col As Integer
    
    Dim OneDayBefore_dateKey As Date
    

    dayCol_start = 4
    hidden_ref_value_col = 37
    MonthValue = "F3"
    Item_col = 3
    FAB_col = 1
    Nick_name_col = 2

    Dim today As Date
    ' ������Ѫ����
    today = Date
    'today = CDate("09/28/2024") 'For testing
    
    ' �����e�u�@��
    Set ws = ThisWorkbook.Sheets(sheetName)
    
    ws.Rows("5:" & ws.Rows.count).ClearContents
    ws.Rows("5:" & ws.Rows.count).Interior.ColorIndex = xlNone ' �M���C��

    ' for monitoring feature
    ws.Cells(5, hidden_ref_value_col).Value = OSAT
    ws.Cells(5, hidden_ref_value_col).Font.Color = RGB(255, 255, 255) ' font white
    
    ' ��� H3 ������L�o��
    monthFilter = ws.Range(MonthValue).Value
    
    ' �q�� 5 ��}�l���L
    rowNum = 5
    
    ' �M���r��
    For Each key In dict.Keys
        Set subDict = dict(key)
        
        
        ' Debug.Print sheetName
        ' Debug.Print key
        
        
        prevdateKey = ""
        no_month_limit_prevdateKey = ""
        
        ' ��l�Ƽg�J���A
        Dim writeToSheet As Boolean
        writeToSheet = False
        
        ' �M���l�r�夤���C�Ӥ��
        For Each dateKey In subDict.Keys
            ' �ˬd dateKey �O�_�����Ĥ��
            If IsDate(dateKey) Then
                ' ����������B��
                targetMonth = Month(CDate(dateKey))
                targetDay = Day(CDate(dateKey))
                
                ' �p�G����P H3 ��������ǰt
                
                'Debug.Print key
                'Debug.Print dateKey
                'Debug.Print subDict(dateKey)(3)
                If key = "SC2000CM3" Then
                    Debug.Print key
                    Debug.Print dateKey
                End If
                
                If targetMonth = monthFilter And (subDict(dateKey)(3) <> "" And subDict(dateKey)(3) <> 0) Then
                    ' ���L E �� key�AA �欰�����W��
                    ' ws.Cells(rowNum, 1).Value = OSAT
                    'ws.Cells(5, 40).Value = OSAT
                    'ws.Cells(5, 40).Font.Color = RGB(255, 255, 255) ' font white
                    
                    ws.Cells(rowNum, Item_col).Value = key
                    
                    ' B, C, E, F ��ƭ�
                    ws.Cells(rowNum, FAB_col).Value = subDict(dateKey)(0) ' B column value
                    ws.Cells(rowNum, Nick_name_col).Value = subDict(dateKey)(1) ' C column value
                    
                    ' �ھڤ���p�� G �檺�����C�� (1 ���q G �}�l�A31 ������ GAK)
                    dayCol = dayCol_start + targetDay ' 8 �O G ���m�AtargetDay �O�Ѽ�
                    
                    ' �b�����Ѽƪ��C��g F �檺��
                    ws.Cells(rowNum, dayCol).Value = subDict(dateKey)(3) ' F column value
                    
                    ' �]�m�g�J�лx�� True�A��ܸӦ榳�ƾڼg�J
                    writeToSheet = True

                    
                    ' �����e�ȻP�e�@�ӭ�
                    If prevdateKey <> "" And CDate(dateKey) = today Then
                        If subDict(dateKey)(4) <> subDict(prevdateKey)(4) Or _
                           subDict(dateKey)(5) <> subDict(prevdateKey)(5) Or _
                           subDict(dateKey)(6) <> subDict(prevdateKey)(6) Then
                            ws.Cells(rowNum, Item_col).Interior.Color = RGB(255, 0, 0) ' �H����
                        End If
                    End If
                    
                    If dayCol = dayCol_start + 1 Then
                        If no_month_limit_prevdateKey <> "" Then
                            If subDict(dateKey)(3) > subDict(no_month_limit_prevdateKey)(3) Then
                                ws.Cells(rowNum, dayCol).Interior.Color = RGB(144, 238, 144) 'light green
                            Else
                                If subDict(dateKey)(3) < subDict(no_month_limit_prevdateKey)(3) Then
                                    ws.Cells(rowNum, dayCol).Interior.Color = RGB(255, 182, 193) 'light red
                                End If
                            End If
                        End If
                    Else
                        If prevdateKey <> "" Then
                            If subDict(dateKey)(3) > subDict(prevdateKey)(3) Then
                                ws.Cells(rowNum, dayCol).Interior.Color = RGB(144, 238, 144) 'light green
                            Else
                                If subDict(dateKey)(3) < subDict(prevdateKey)(3) Then
                                    ws.Cells(rowNum, dayCol).Interior.Color = RGB(255, 182, 193) 'light red
                                End If
                            End If
                        End If
                        
                        OneDayBefore_dateKey = DateAdd("d", -1, CDate(dateKey)) ' �p��e�@�Ѫ���
                        Dim OneDayBefore_dateKeyStr As String
                        OneDayBefore_dateKeyStr = Format(OneDayBefore_dateKey, "m/d/yyyy")
                        If prevdateKey <> OneDayBefore_dateKeyStr Then
                            ws.Cells(rowNum, dayCol).Interior.Color = RGB(144, 238, 144) 'light green
                        End If
                    End If
                    
                    
                    prevdateKey = dateKey
                End If
                
                
                
                no_month_limit_prevdateKey = dateKey
                
            Else
                ' �p�G���O���Ī�����A���L�o�@���O��
                Debug.Print "Invalid date format for key: " & dateKey
            End If
            
        Next dateKey
        
        ' �p�G���ƾڼg�J�A�h�W�[�渹
        If writeToSheet Then
            ' �եζ�R�ŭȪ����
            ' FillEmptyCells ws, rowNum, 8, dayCol ' 8 �O G ��AdayCol �O�����C
            rowNum = rowNum + 1
        End If
        


    Next key
End Function




