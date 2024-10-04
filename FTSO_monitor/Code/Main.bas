Attribute VB_Name = "Main"
Dim dictPTI As Object
Dim dictASE As Object
Dim dictSigurd As Object


Sub ProcessSheet(sheetName As String)
    Dim dict As Object
    
    ' 根據工作表名稱選擇對應的字典
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
    
    ' 打印字典到相應的工作表
    PrintDictionaryToSheet dict, sheetName, "SO Summary"
End Sub



Sub CreateDictionariesForAllSheets()
    ' 創建並存儲 PTI, ASE, Sigurd 三個分頁的字典
    Set dictPTI = CreateDictionaryForSheet("PTI")
    Set dictASE = CreateDictionaryForSheet("ASE")
    Set dictSigurd = CreateDictionaryForSheet("Sigurd")
    
    ' 調試輸出，確認字典已創建
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
    
    ' 設置工作表
    Set ws = ThisWorkbook.Sheets(sheetName)
    
    ' 創建主字典
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' 找到最後一行
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    
    ' 創建日期集合
    Set dateList = New Collection
    
    ' 遍歷每一行，從第二行開始
    For rowNum = 2 To lastRow
        dateKey = ws.Cells(rowNum, 1).Value ' A 欄的值作為日期 key
        
        ' 將日期加入集合中（避免重複）
        On Error Resume Next
        dateList.Add dateKey, CStr(dateKey)
        On Error GoTo 0
    Next rowNum
    
    ' 將日期集合轉換為數組並排序
    sortedDates = SortCollection(dateList)

    ' 遍歷排序後的日期
    For i = LBound(sortedDates) To UBound(sortedDates)
        dateKey = sortedDates(i)
        
        ' 再次遍歷每一行以填充字典
        For rowNum = 2 To lastRow
            key = ws.Cells(rowNum, 4).Value ' D 欄的值作為主字典的 key
            
            ' 只處理符合當前日期的行
            If ws.Cells(rowNum, 1).Value = dateKey Then
                bValue = ws.Cells(rowNum, 2).Value ' B 欄的值
                cValue = ws.Cells(rowNum, 3).Value ' C 欄的值
                eValue = ws.Cells(rowNum, 5).Value ' E 欄的值
                fValue = ws.Cells(rowNum, 6).Value ' F 欄的值
                
                gValue = ws.Cells(rowNum, 7).Value ' G 欄的值
                hValue = ws.Cells(rowNum, 8).Value ' H 欄的值
                iValue = ws.Cells(rowNum, 9).Value ' I 欄的值
                
                ' 如果主字典中不存在該 key，則創建新的子字典
                If Not dict.Exists(key) Then
                    Set subDict = CreateObject("Scripting.Dictionary")
                    dict.Add key, subDict
                Else
                    Set subDict = dict(key) ' 如果 key 已經存在，取出子字典
                End If
                
                ' 檢查子字典中是否已有日期作為 key
                If Not subDict.Exists(dateKey) Then
                    ' 將 B, C, E, F 欄的值存入數據列表
                    subDict.Add dateKey, Array(bValue, cValue, eValue, fValue, gValue, hValue, iValue)
                Else
                    ' 可選：處理相同日期的情況，這裡不處理重複日期
                End If
            End If
        Next rowNum
    Next i
    
    ' 返回創建的字典
    Set CreateDictionaryForSheet = dict
End Function

Function SortCollection(col As Collection) As Variant
    Dim arr() As Variant
    Dim i As Long, j As Long
    Dim temp As Variant

    ' 將集合轉換為數組
    ReDim arr(1 To col.count)
    For i = 1 To col.count
        arr(i) = col(i)
    Next i

    ' 冒泡排序
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
    ' 獲取今天的日期
    today = Date
    'today = CDate("09/28/2024") 'For testing
    
    ' 獲取當前工作表
    Set ws = ThisWorkbook.Sheets(sheetName)
    
    ws.Rows("5:" & ws.Rows.count).ClearContents
    ws.Rows("5:" & ws.Rows.count).Interior.ColorIndex = xlNone ' 清除顏色

    ' for monitoring feature
    ws.Cells(5, hidden_ref_value_col).Value = OSAT
    ws.Cells(5, hidden_ref_value_col).Font.Color = RGB(255, 255, 255) ' font white
    
    ' 獲取 H3 的月份過濾值
    monthFilter = ws.Range(MonthValue).Value
    
    ' 從第 5 行開始打印
    rowNum = 5
    
    ' 遍歷字典
    For Each key In dict.Keys
        Set subDict = dict(key)
        
        
        ' Debug.Print sheetName
        ' Debug.Print key
        
        
        prevdateKey = ""
        no_month_limit_prevdateKey = ""
        
        ' 初始化寫入狀態
        Dim writeToSheet As Boolean
        writeToSheet = False
        
        ' 遍歷子字典中的每個日期
        For Each dateKey In subDict.Keys
            ' 檢查 dateKey 是否為有效日期
            If IsDate(dateKey) Then
                ' 拆分日期成月、日
                targetMonth = Month(CDate(dateKey))
                targetDay = Day(CDate(dateKey))
                
                ' 如果月份與 H3 中的月份匹配
                
                'Debug.Print key
                'Debug.Print dateKey
                'Debug.Print subDict(dateKey)(3)
                If key = "SC2000CM3" Then
                    Debug.Print key
                    Debug.Print dateKey
                End If
                
                If targetMonth = monthFilter And (subDict(dateKey)(3) <> "" And subDict(dateKey)(3) <> 0) Then
                    ' 打印 E 欄 key，A 欄為分頁名稱
                    ' ws.Cells(rowNum, 1).Value = OSAT
                    'ws.Cells(5, 40).Value = OSAT
                    'ws.Cells(5, 40).Font.Color = RGB(255, 255, 255) ' font white
                    
                    ws.Cells(rowNum, Item_col).Value = key
                    
                    ' B, C, E, F 欄數值
                    ws.Cells(rowNum, FAB_col).Value = subDict(dateKey)(0) ' B column value
                    ws.Cells(rowNum, Nick_name_col).Value = subDict(dateKey)(1) ' C column value
                    
                    ' 根據日期計算 G 欄的偏移列數 (1 號從 G 開始，31 號對應 GAK)
                    dayCol = dayCol_start + targetDay ' 8 是 G 欄位置，targetDay 是天數
                    
                    ' 在對應天數的列填寫 F 欄的值
                    ws.Cells(rowNum, dayCol).Value = subDict(dateKey)(3) ' F column value
                    
                    ' 設置寫入標誌為 True，表示該行有數據寫入
                    writeToSheet = True

                    
                    ' 比較當前值與前一個值
                    If prevdateKey <> "" And CDate(dateKey) = today Then
                        If subDict(dateKey)(4) <> subDict(prevdateKey)(4) Or _
                           subDict(dateKey)(5) <> subDict(prevdateKey)(5) Or _
                           subDict(dateKey)(6) <> subDict(prevdateKey)(6) Then
                            ws.Cells(rowNum, Item_col).Interior.Color = RGB(255, 0, 0) ' 淡紅色
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
                        
                        OneDayBefore_dateKey = DateAdd("d", -1, CDate(dateKey)) ' 計算前一天的切
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
                ' 如果不是有效的日期，跳過這一條記錄
                Debug.Print "Invalid date format for key: " & dateKey
            End If
            
        Next dateKey
        
        ' 如果有數據寫入，則增加行號
        If writeToSheet Then
            ' 調用填充空值的函數
            ' FillEmptyCells ws, rowNum, 8, dayCol ' 8 是 G 欄，dayCol 是結束列
            rowNum = rowNum + 1
        End If
        


    Next key
End Function




