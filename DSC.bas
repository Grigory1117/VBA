Attribute VB_Name = "DSC"
' ╔══════════════════════════════════════════════════════════════════╗
' ║                      主程序：一次登入爬取所有網站                      ║
' ╚══════════════════════════════════════════════════════════════════╝

Sub Final()
    ' ========== 只輸入一次帳號密碼 ==========
    Dim username As String, password As String
    Dim yearCurrent As String, monthCurrent As String

    username = InputBox("請輸入帳號:", "登入")
    If username = "" Then Exit Sub

    password = InputBox("請輸入密碼:", "登入")
    If password = "" Then Exit Sub

    yearCurrent = InputBox("請輸入年份:", "查詢期間", "2026")
    If yearCurrent = "" Then Exit Sub
    
    monthCurrent = InputBox("請輸入月份:", "查詢期間", "12")
    If monthCurrent = "" Then Exit Sub
    
    If Len(monthCurrent) = 1 Then monthCurrent = "0" & monthCurrent
    
    ' ========== 創建 IE  ==========
    Dim IE As Object
    Set IE = CreateObject("InternetExplorer.Application")
    IE.Visible = True
    
    ' 登入（只需一次）
    IE.navigate "https://example?"
    Do While IE.Busy Or IE.readyState <> 4: DoEvents: Loop
    
    IE.document.getElementById("uxUserId").Value = username
    IE.document.getElementById("uxPassword").Value = password
    IE.document.getElementById("uxSubmit").Click
    
    Do While IE.Busy Or IE.readyState <> 4: DoEvents: Loop
    Application.Wait Now + timeValue("00:00:03")
    
    ' ========== 第一站：example5 ==========
    Debug.Print "==============================================="
    Debug.Print "========== 開始爬取 example5 =========="
    Debug.Print "==============================================="
    
    Call Crawlexample5(IE, yearCurrent, monthCurrent)
    
    ' ========== 第二站：example1 ==========
    Debug.Print "==============================================="
    Debug.Print "========== 開始爬取 example1 =========="
    Debug.Print "==============================================="
    
    Call Crawlexample1(IE, yearCurrent, monthCurrent)
    
    ' ========== 第三站：exampleR7 ==========
    Debug.Print "==============================================="
    Debug.Print "========== 開始爬取 exampleR7 =========="
    Debug.Print "==============================================="
    
    Call CrawlexampleR7(IE)
    
    ' ========== 完成，關閉 IE ==========
    IE.Quit
    
    ' ========== 執行資料更新 ==========
    Debug.Print "==============================================="
    Debug.Print "========== 執行資料更新 =========="
    Debug.Print "==============================================="
    
    Call UpdateMonthlyToYearly
    
    ' ========== 全部完成 ==========
    MsgBox "所有資料已成功匯入並更新！" & vbCrLf & vbCrLf & _
           "example5 已完成" & vbCrLf & _
           "example1 已完成" & vbCrLf & _
           "exampleR7 已完成" & vbCrLf & _
           "資料更新已完成" & vbCrLf & vbCrLf & _
           "期間：" & yearCurrent & "/" & monthCurrent, vbInformation, "爬取完成"
End Sub


' ╔══════════════════════════════════════════════════════════════════╗
' ║                          爬蟲子函數區                              ║
' ╚══════════════════════════════════════════════════════════════════╝

' ========== 子函數：爬取 example5 ==========
Sub Crawlexample5(IE As Object, yearCurrent As String, monthCurrent As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("example5")
    ws.Cells.Clear
    
    IE.navigate "https://example5?"
    Do While IE.Busy Or IE.readyState <> 4: DoEvents: Loop
    Application.Wait Now + timeValue("00:00:03")
    
    Call ExpandSecondQueryPanel(IE)
    
    Debug.Print "========== 填入 example5 查詢條件 =========="
    On Error Resume Next
    
    ' 第一組（單月）
    Call SetVueData(IE, "queryParam1.yearFm", yearCurrent)
    Call SetVueData(IE, "queryParam1.verSrlFm", monthCurrent)
    Call SetVueData(IE, "queryParam1.yearTo", yearCurrent)
    Call SetVueData(IE, "queryParam1.verSrlTo", monthCurrent)
    
    Debug.Print "  單月條件: " & yearCurrent & "/" & monthCurrent
    
    ' 第二組（年累）
    Call SetVueData(IE, "queryParam2.compId", "0000")
    Call SetVueData(IE, "queryParam2.ver", "PD")
    Call SetVueData(IE, "queryParam2.book", "A")
    Call SetVueData(IE, "queryParam2.rcdCode", "ACT1,ACT6")
    Call SetVueData(IE, "queryParam2.yearFm", yearCurrent)
    Call SetVueData(IE, "queryParam2.verSrlFm", "01")
    Call SetVueData(IE, "queryParam2.yearTo", yearCurrent)
    Call SetVueData(IE, "queryParam2.verSrlTo", monthCurrent)
    
    Debug.Print "  年累條件: " & yearCurrent & "/01 ~ " & yearCurrent & "/" & monthCurrent
    
    ' 點擊查詢
    Debug.Print "========== 查詢 example5 =========="
    Dim btnQuery As Object
    Set btnQuery = IE.document.getElementById("btnQuery")
    If Not btnQuery Is Nothing Then
        btnQuery.Click
    Else
        IE.document.parentWindow.execScript "document.getElementById('btnQuery').click();", "JavaScript"
    End If
    
    WaitForIE IE
    Application.Wait Now + timeValue("00:00:05")
    
    ' 提取數據
    Debug.Print "========== 提取 example5 數據 =========="
    Dim panels As Object
    Set panels = IE.document.getElementsByClassName("panel-primary")
    
    If panels.Length > 0 Then
        Debug.Print "  找到 " & panels.Length & " 個 panel"
        
        If panels.Length >= 1 Then
            Dim tables1 As Object
            Set tables1 = panels(0).getElementsByTagName("table")
            
            If tables1.Length > 0 Then
                Call ExtractTableData(tables1(0), ws, 1, 1)
                
                ' example5 特定的資料處理
                ws.Range("A2:E2").Insert Shift:=xlToRight
                ws.Range("G1:J1").Insert Shift:=xlToRight
                ws.Range("L1:O1").Insert Shift:=xlToRight
                ws.Cells(1, 6).Font.Bold = True
                ws.Cells(1, 6).Font.Color = RGB(255, 0, 0)
                ws.Cells(1, 11).Font.Bold = True
                ws.Cells(1, 11).Font.Color = RGB(255, 0, 0)
                
                Debug.Print "example5 數據提取成功"
            End If
        End If
    Else
        Debug.Print "example5 找不到資料面板"
    End If
    
    On Error GoTo 0
End Sub

' ========== 子函數：爬取 example1 ==========
Sub Crawlexample1(IE As Object, yearCurrent As String, monthCurrent As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("example1")
    ws.Cells.Clear
    
    IE.navigate "https://example1?"
    Do While IE.Busy Or IE.readyState <> 4: DoEvents: Loop
    Application.Wait Now + timeValue("00:00:03")
    
    Call ExpandSecondQueryPanel(IE)
    
    Debug.Print "========== 填入 example1 查詢條件 =========="
    On Error Resume Next
    
    ' 第一組（單月） - example1 特有的 alloc 參數
    Call SetVueData(IE, "queryParam1.alloc", "8P0")
    Call SetVueData(IE, "queryParam1.yearFm", yearCurrent)
    Call SetVueData(IE, "queryParam1.verSrlFm", monthCurrent)
    Call SetVueData(IE, "queryParam1.yearTo", yearCurrent)
    Call SetVueData(IE, "queryParam1.verSrlTo", monthCurrent)
    
    Debug.Print "  單月條件: " & yearCurrent & "/" & monthCurrent & " (alloc=8P0)"
    
    ' 第二組（年累）
    Call SetVueData(IE, "queryParam2.compId", "0000")
    Call SetVueData(IE, "queryParam2.ver", "PD")
    Call SetVueData(IE, "queryParam2.book", "A")
    Call SetVueData(IE, "queryParam2.rcdCode", "ACT1,ACT6")
    Call SetVueData(IE, "queryParam2.alloc", "8P0")
    Call SetVueData(IE, "queryParam2.yearFm", yearCurrent)
    Call SetVueData(IE, "queryParam2.verSrlFm", "01")
    Call SetVueData(IE, "queryParam2.yearTo", yearCurrent)
    Call SetVueData(IE, "queryParam2.verSrlTo", monthCurrent)
    
    Debug.Print "  年累條件: " & yearCurrent & "/01 ~ " & yearCurrent & "/" & monthCurrent & " (alloc=8P0)"
    
    ' 點擊查詢
    Debug.Print "========== 查詢 example1 =========="
    Dim btnQuery As Object
    Set btnQuery = IE.document.getElementById("btnQuery")
    If Not btnQuery Is Nothing Then
        btnQuery.Click
    Else
        IE.document.parentWindow.execScript "document.getElementById('btnQuery').click();", "JavaScript"
    End If
    
    WaitForIE IE
    Application.Wait Now + timeValue("00:00:05")
    
    ' 提取數據
    Debug.Print "========== 提取 example1 數據 =========="
    Dim panels As Object
    Set panels = IE.document.getElementsByClassName("panel-primary")
    
    If panels.Length > 0 Then
        Debug.Print "  找到 " & panels.Length & " 個 panel"
        
        If panels.Length >= 1 Then
            Dim tables1 As Object
            Set tables1 = panels(0).getElementsByTagName("table")
            
            If tables1.Length > 0 Then
                Call ExtractTableData(tables1(0), ws, 1, 1)
                
                ' example1 特定的插入儲存格
                ws.Range("A2:G2").Insert Shift:=xlToRight
                ws.Range("I1:J1").Insert Shift:=xlToRight
                ws.Range("L1:M1").Insert Shift:=xlToRight
                ws.Cells(1, 8).Font.Bold = True
                ws.Cells(1, 8).Font.Color = RGB(255, 0, 0)
                ws.Cells(1, 11).Font.Bold = True
                ws.Cells(1, 11).Font.Color = RGB(255, 0, 0)
                
                Debug.Print "example1 數據提取成功"
            End If
        End If
    Else
        Debug.Print "example1 找不到資料面板"
    End If
    
    On Error GoTo 0
End Sub

' ========== 子函數：爬取 exampleR7（含自動下載） ==========
Sub CrawlexampleR7(IE As Object)
    IE.navigate "https://exampleR7"
    
    Do While IE.Busy Or IE.readyState <> 4: DoEvents: Loop
    Application.Wait Now + timeValue("00:00:03")
    
    Debug.Print "========== 開始查詢 exampleR7 =========="
    
    Dim btnQuery As Object
    Set btnQuery = IE.document.getElementById("queryButton")
    If Not btnQuery Is Nothing Then
        btnQuery.Click
    Else
        IE.document.parentWindow.execScript "document.getElementById('queryButton').click();", "JavaScript"
    End If
    
    WaitForIE IE
    Application.Wait Now + timeValue("00:00:05")
    
    Debug.Print "========== 抓取流水碼 =========="
    
    Dim panels As Object
    Set panels = IE.document.getElementsByClassName("col-sm-12")
    
    If panels.Length > 0 Then
        Dim tables As Object
        Set tables = panels(0).getElementsByTagName("table")
        
        If tables.Length > 0 Then
            Dim tbody As Object
            Set tbody = tables(0).getElementsByTagName("tbody")
            
            If tbody.Length > 0 Then
                Dim rows As Object
                Set rows = tbody(0).getElementsByTagName("tr")
                
                If rows.Length > 0 Then
                    Dim serialCode As String
                    Dim firstRow As Object
                    Set firstRow = rows(0)
                    
                    serialCode = GetSerialCodeFromRow(firstRow)
                    
                    If serialCode <> "" And Len(serialCode) = 14 Then
                        Debug.Print "最終流水碼: " & serialCode
                        
                        Call ProcessAexampleR7WithDownload(serialCode, IE, firstRow)
                        
                        Debug.Print "exampleR7 處理完成"
                        Exit Sub
                    Else
                        MsgBox "流水碼格式錯誤！" & vbCrLf & "取得: " & serialCode, vbExclamation
                    End If
                Else
                    Debug.Print "表格中沒有數據"
                End If
            End If
        End If
    End If
End Sub


' ╔══════════════════════════════════════════════════════════════════╗
' ║                         Vue 操作函數區                             ║
' ╚══════════════════════════════════════════════════════════════════╝

Sub ExpandSecondQueryPanel(IE As Object)
    Dim js As String
    js = "var panel=document.getElementById('qryMethodA');"
    js = js & "if(panel){panel.classList.remove('collapse');panel.classList.add('in');panel.style.display='block';}"
    IE.document.parentWindow.execScript js, "JavaScript"
    Debug.Print "已展開第二個查詢區塊"
End Sub

Sub SetVueData(IE As Object, dataPath As String, Value As String)
    Dim js As String
    
    js = "var app=document.querySelector('#app').__vue__;"
    js = js & "var path='" & dataPath & "'.split('.');"
    js = js & "var obj=app;"
    js = js & "for(var i=0;i<path.length-1;i++){obj=obj[path[i]];}"
    js = js & "obj[path[path.length-1]]='" & Value & "';"
    js = js & "app.$forceUpdate();"
    
    IE.document.parentWindow.execScript js, "JavaScript"
    Application.Wait Now + timeValue("00:00:01")
    
    js = "var elem=document.querySelector('[v-model=""" & dataPath & """]');"
    js = js & "if(elem){elem.value='" & Value & "';elem.dispatchEvent(new Event('change',{bubbles:true}));}"
    
    IE.document.parentWindow.execScript js, "JavaScript"
    Debug.Print "    已設定: " & dataPath & " = " & Value
End Sub


' ╔══════════════════════════════════════════════════════════════════╗
' ║                       資料提取與處理函數區                          ║
' ╚══════════════════════════════════════════════════════════════════╝

Sub ExtractTableData(table As Object, ws As Worksheet, startRow As Long, startCol As Long)
    Dim rows As Object
    Set rows = table.rows
    
    Dim row As Long, col As Long
    Dim Cells As Object
    
    For row = 0 To rows.Length - 1
        Set Cells = rows(row).Cells
        For col = 0 To Cells.Length - 1
            ws.Cells(startRow + row, startCol + col).Value = Trim(Cells(col).innerText)
        Next col
    Next row
End Sub

Function GetSerialCodeFromRow(row As Object) As String
    On Error Resume Next
    
    Dim Cells As Object
    Set Cells = row.getElementsByTagName("td")
    
    If Cells.Length >= 4 Then
        Dim dateCell As Object
        Dim timeCell As Object
        Set dateCell = Cells(2)
        Set timeCell = Cells(3)
        
        Dim rawText As String
        Dim cleanText As String
        
        rawText = Trim(dateCell.innerText) & Trim(timeCell.innerText)
        Debug.Print "    原始文字: " & rawText
        
        cleanText = ""
        Dim i As Long
        For i = 1 To Len(rawText)
            If Mid(rawText, i, 1) >= "0" And Mid(rawText, i, 1) <= "9" Then
                cleanText = cleanText & Mid(rawText, i, 1)
            End If
        Next i
        
        Debug.Print "    清理後: " & cleanText
        Debug.Print "    長度: " & Len(cleanText)
        
        Dim dateValue As String
        Dim timeValue As String
        
        If Len(cleanText) >= 30 Then
            dateValue = Left(cleanText, 8)
            timeValue = Mid(cleanText, 25, 6)
            
            Debug.Print "    提取日期: " & dateValue
            Debug.Print "    提取時間: " & timeValue
            
            GetSerialCodeFromRow = dateValue & timeValue
            Debug.Print "    流水碼: " & GetSerialCodeFromRow
        Else
            GetSerialCodeFromRow = ""
            Debug.Print "  找不到流水碼"
        End If
    Else
        GetSerialCodeFromRow = ""
    End If
    
    On Error GoTo 0
End Function

Sub ProcessAexampleR7WithDownload(serialNo As String, IE As Object, row As Object)
    Debug.Print "========== 處理 exampleR7 檔案 =========="
    Debug.Print "  流水碼: " & serialNo
    
    Dim fileName As String
    Dim filePath As String
    Dim downloadFolder As String
    Dim currentUser As String
    Dim wbSource As Workbook
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    
    currentUser = Environ("USERNAME")
    Set wsTarget = ThisWorkbook.Sheets("exampleR7")
    
    ' 設定下載路徑
    downloadFolder = "C:\Users\" & currentUser & "\Downloads\"
    fileName = "exampleR7." & serialNo & ".XLS"
    filePath = downloadFolder & fileName
    
    Debug.Print "  Windows 使用者: " & currentUser
    Debug.Print "  下載路徑: " & downloadFolder
    Debug.Print "  檔案路徑: " & filePath
    
    ' ========== 檢查檔案是否存在 ==========
    If Dir(filePath) = "" Then
        Debug.Print "  檔案不存在，開始自動下載..."
        
        ' 改用自動處理 IE 下載對話框的方法
        Dim downloadSuccess As Boolean
        downloadSuccess = DownloadFileAutoClick(IE, row, downloadFolder, fileName)
        
        If Not downloadSuccess Then
            MsgBox "下載失敗！請手動下載後重試。", vbExclamation
            Exit Sub
        End If
        
        Debug.Print "  下載完成"
    Else
        Debug.Print "  檔案已存在"
    End If
    
    ' ========== 開啟並複製檔案 ==========
    If Dir(filePath) <> "" Then
        On Error Resume Next
        Set wbSource = Workbooks.Open(filePath)
        
        If Err.Number <> 0 Then
            MsgBox "開啟檔案失敗：" & Err.Description, vbCritical
            Exit Sub
        End If
        On Error GoTo 0

        Set wsSource = wbSource.Sheets("依廠別彙總")
        
        Dim lastRow As Long
        Dim lastCol As Long
        Dim startRow As Long
        lastRow = wsSource.Cells(wsSource.rows.Count, 1).End(xlUp).row
        lastCol = wsSource.Cells(1, wsSource.Columns.Count).End(xlToLeft).Column '31
        
        wsTarget.Range("B13").Resize(lastRow, 31).Value = _
        wsSource.Range(wsSource.Cells(1, 1), _
        wsSource.Cells(lastRow, 31)).Value

        wbSource.Close SaveChanges:=False
        Application.CutCopyMode = False
        startRow = 15
        '逐行填入公式
        For i = startRow To startRow + lastRow - 1
            wsTarget.Cells(i, "A").Formula = "=RIGHT($F$13,2)&RIGHT($F$14,1)&D" & i & "&E" & i
        Next i
            

        Debug.Print "  檔案處理完成"
    Else
        MsgBox "下載後仍找不到檔案！", vbExclamation
    End If
End Sub

' ========== 新增：自動處理 IE 下載對話框 ==========
Function DownloadFileAutoClick(IE As Object, row As Object, downloadFolder As String, expectedFileName As String) As Boolean
    On Error Resume Next
    
    Debug.Print "========== 開始自動下載 =========="
    
    ' 找到下載按鈕
    Dim Cells As Object
    Set Cells = row.getElementsByTagName("td")
    
    Dim downloadCell As Object
    Set downloadCell = Cells(Cells.Length - 2)
    
    Dim buttons As Object
    Set buttons = downloadCell.getElementsByTagName("button")
    
    Debug.Print "  找到 " & buttons.Length & " 個按鈕"
    
    Dim downloadBtn As Object
    Dim i As Long
    
    For i = buttons.Length - 1 To 0 Step -1
        If InStr(buttons(i).innerText, "下載原始檔") > 0 Then
            Set downloadBtn = buttons(i)
            Debug.Print "  找到「下載原始檔」按鈕"
            Exit For
        End If
    Next i
    
    If downloadBtn Is Nothing Then
        For i = buttons.Length - 1 To 0 Step -1
            If buttons(i).Style.display <> "none" Then
                Set downloadBtn = buttons(i)
                Debug.Print "  找到最後一個可見按鈕: " & buttons(i).innerText
                Exit For
            End If
        Next i
    End If
    
    If downloadBtn Is Nothing Then
        Debug.Print "  找不到下載按鈕"
        DownloadFileAutoClick = False
        Exit Function
    End If
    
    ' 記錄下載前的狀態
    Dim fileExistedBefore As Boolean
    fileExistedBefore = (Dir(downloadFolder & expectedFileName) <> "")
    
    ' 點擊下載按鈕
    Debug.Print "  點擊下載按鈕..."
    downloadBtn.Click
    
    ' 等待 IE 下載對話框出現（增加等待時間）
    Debug.Print "  等待下載對話框出現..."
    Application.Wait Now + timeValue("00:00:03")  ' 改成3秒
    
    ' 嘗試多種按鍵組合
    Debug.Print "  嘗試自動點擊儲存按鈕..."
    
    ' 方法1：直接按 Alt+S（儲存）
    Debug.Print "    方法1: Alt+S"
    Application.SendKeys "%s", True
    Application.Wait Now + timeValue("00:00:01")
    
    ' 檢查是否成功，如果沒有則嘗試下一個方法
    If Dir(downloadFolder & expectedFileName) = "" Then
        Debug.Print "    方法1 似乎沒效果，嘗試方法2..."
        
        ' 方法2：Tab 到儲存按鈕後按 Enter
        Debug.Print "    方法2: Tab + Enter"
        Application.SendKeys "{TAB}", True
        Application.Wait Now + timeValue("00:00:01")
        Application.SendKeys "{ENTER}", True
        Application.Wait Now + timeValue("00:00:01")
    End If
    
    ' 再檢查一次
    If Dir(downloadFolder & expectedFileName) = "" Then
        Debug.Print "    方法2 似乎也沒效果，嘗試方法3..."
        
        ' 方法3：直接按 Enter（如果儲存是預設按鈕）
        Debug.Print "    方法3: 直接 Enter"
        Application.SendKeys "{ENTER}", True
        Application.Wait Now + timeValue("00:00:01")
    End If
    
    ' 再檢查一次
    If Dir(downloadFolder & expectedFileName) = "" Then
        Debug.Print "    方法3 似乎也沒效果，嘗試方法4..."
        
        ' 方法4：Space 鍵（按下目前焦點的按鈕）
        Debug.Print "    方法4: Space"
        Application.SendKeys " ", True
        Application.Wait Now + timeValue("00:00:01")
    End If
    
    ' 等待下載開始
    Debug.Print "  等待下載開始..."
    Application.Wait Now + timeValue("00:00:02")
    
    ' 等待檔案下載完成
    Dim waitCount As Integer
    Dim fileDownloaded As Boolean
    waitCount = 0
    fileDownloaded = False
    
    Debug.Print "  等待下載完成..."
    
    Do While waitCount < 60  ' 最多等60秒
        Application.Wait Now + timeValue("00:00:01")
        waitCount = waitCount + 1
        DoEvents
        
        ' 檢查檔案是否存在
        If Dir(downloadFolder & expectedFileName) <> "" Then
            ' 檢查檔案大小是否穩定（確保下載完成）
            Dim fileSize1 As Long, fileSize2 As Long
            fileSize1 = FileLen(downloadFolder & expectedFileName)
            
            Application.Wait Now + timeValue("00:00:02")
            
            fileSize2 = FileLen(downloadFolder & expectedFileName)
            
            ' 如果檔案大小穩定且大於 0
            If fileSize1 = fileSize2 And fileSize1 > 0 Then
                fileDownloaded = True
                Debug.Print "  檔案已下載（等待 " & waitCount & " 秒，大小: " & fileSize1 & " bytes）"
                Exit Do
            End If
        End If
        
        ' 每5秒顯示一次進度
        If waitCount Mod 5 = 0 Then
            Debug.Print "  等待中... (" & waitCount & " 秒)"
        End If
    Loop
    
    If Not fileDownloaded Then
        Debug.Print "  下載逾時（等待超過60秒）"
        
        ' 顯示提示訊息，讓使用者手動處理
        MsgBox "自動下載似乎沒有成功。" & vbCrLf & vbCrLf & _
               "請手動點擊下載對話框的「儲存」按鈕，" & vbCrLf & _
               "然後程式會繼續執行。" & vbCrLf & vbCrLf & _
               "檔案應儲存到：" & vbCrLf & downloadFolder, _
               vbExclamation, "需要手動操作"
        
        ' 再等一次（給使用者手動點擊的時間）
        waitCount = 0
        Do While waitCount < 60
            Application.Wait Now + timeValue("00:00:01")
            waitCount = waitCount + 1
            DoEvents
            
            If Dir(downloadFolder & expectedFileName) <> "" Then
                Dim fs1 As Long, fs2 As Long
                fs1 = FileLen(downloadFolder & expectedFileName)
                Application.Wait Now + timeValue("00:00:02")
                fs2 = FileLen(downloadFolder & expectedFileName)
                
                If fs1 = fs2 And fs1 > 0 Then
                    fileDownloaded = True
                    Debug.Print "  手動下載完成"
                    Exit Do
                End If
            End If
        Loop
        
        DownloadFileAutoClick = fileDownloaded
    Else
        DownloadFileAutoClick = True
    End If
    
    On Error GoTo 0
End Function

Function DownloadFileFromRow(IE As Object, row As Object, downloadFolder As String, expectedFileName As String) As Boolean
    On Error Resume Next
    
    Debug.Print "========== 開始下載 =========="
    
    Dim Cells As Object
    Set Cells = row.getElementsByTagName("td")
    
    Dim downloadCell As Object
    Set downloadCell = Cells(Cells.Length - 2)
    
    Dim buttons As Object
    Set buttons = downloadCell.getElementsByTagName("button")
    
    Debug.Print "  找到 " & buttons.Length & " 個按鈕"
    
    Dim downloadBtn As Object
    Dim i As Long
    
    For i = buttons.Length - 1 To 0 Step -1
        If InStr(buttons(i).innerText, "下載原始檔") > 0 Then
            Set downloadBtn = buttons(i)
            Debug.Print "找到「下載原始檔」按鈕"
            Exit For
        End If
    Next i
    
    If downloadBtn Is Nothing Then
        For i = buttons.Length - 1 To 0 Step -1
            If buttons(i).Style.display <> "none" Then
                Set downloadBtn = buttons(i)
                Debug.Print "  找到最後一個可見按鈕: " & buttons(i).innerText
                Exit For
            End If
        Next i
    End If
    
    If downloadBtn Is Nothing Then
        Debug.Print "找不到下載按鈕"
        DownloadFileFromRow = False
        Exit Function
    End If
    
    Dim fileExistedBefore As Boolean
    Dim oldFileTime As Date
    
    fileExistedBefore = (Dir(downloadFolder & expectedFileName) <> "")
    If fileExistedBefore Then
        oldFileTime = FileDateTime(downloadFolder & expectedFileName)
    End If
    
    Debug.Print "  點擊下載按鈕..."
    downloadBtn.Click
    
    Application.Wait Now + timeValue("00:00:02")
    
    Dim waitCount As Integer
    Dim fileDownloaded As Boolean
    waitCount = 0
    fileDownloaded = False
    
    Debug.Print "  等待下載完成..."
    
    Do While waitCount < 30
        Application.Wait Now + timeValue("00:00:01")
        waitCount = waitCount + 1
        
        If Dir(downloadFolder & expectedFileName) <> "" Then
            Dim currentFileTime As Date
            currentFileTime = FileDateTime(downloadFolder & expectedFileName)
            
            If Not fileExistedBefore Or currentFileTime > oldFileTime Then
                Application.Wait Now + timeValue("00:00:02")
                
                If FileLen(downloadFolder & expectedFileName) > 0 Then
                    fileDownloaded = True
                    Debug.Print "檔案已下載（等待 " & waitCount & " 秒）"
                    Exit Do
                End If
            End If
        End If
        
        If waitCount Mod 5 = 0 Then
            Debug.Print "  等待中... (" & waitCount & " 秒)"
        End If
    Loop
    
    If Not fileDownloaded Then
        Debug.Print "下載逾時（等待超過60秒）"
        DownloadFileFromRow = False
    Else
        DownloadFileFromRow = True
    End If
    
    On Error GoTo 0
End Function


' ╔══════════════════════════════════════════════════════════════════╗
' ║                         資料更新函數區                             ║
' ╚══════════════════════════════════════════════════════════════════╝

Sub UpdateMonthlyToYearly()
    Dim wsMonthly As Worksheet
    Dim wsYearly As Worksheet
    Dim lastRowMonthly As Long
    Dim lastRowYearly As Long
    Dim i As Long, r As Long
    Dim monthStr As String
    Dim salesType As String
    Dim itemName As String
    Dim isFound As Boolean
    
    Set wsMonthly = Worksheets("exampleR7")
    Set wsYearly = Worksheets("明細")
    
    lastRowMonthly = wsMonthly.Cells(11, 3).End(xlUp).row
    lastRowYearly = wsYearly.Cells(wsYearly.rows.Count, 2).End(xlUp).row
    
    Debug.Print "========== 更新年累積資料 =========="
    Debug.Print "  月報表資料列數: " & lastRowMonthly
    Debug.Print "  年累表資料列數: " & lastRowYearly
    
    For i = 2 To lastRowMonthly
        monthStr = Trim(wsMonthly.Cells(i, 3).Value)
        salesType = Trim(wsMonthly.Cells(i, 4).Value)
        itemName = Trim(wsMonthly.Cells(i, 5).Value)
        
        isFound = False
        
        For r = 2 To lastRowYearly
            If Trim(wsYearly.Cells(r, 2).Value) = monthStr _
            And Trim(wsYearly.Cells(r, 3).Value) = salesType _
            And Trim(wsYearly.Cells(r, 4).Value) = itemName Then
                
                wsYearly.Cells(r, 5).Resize(1, 7).Value = _
                wsMonthly.Cells(i, 6).Resize(1, 7).Value
                
                isFound = True
                Exit For
            End If
        Next r
        
        If Not isFound Then
            Debug.Print "  找不到：" & monthStr & " / " & salesType & " / " & itemName
        End If
    Next i
    
    Debug.Print "全部更新完成"
End Sub


' ╔══════════════════════════════════════════════════════════════════╗
' ║                         通用輔助函數區                             ║
' ╚══════════════════════════════════════════════════════════════════╝

Sub WaitForIE(IE As Object)
    Do While IE.Busy Or IE.readyState <> 4
        DoEvents
    Loop
    Application.Wait Now + timeValue("00:00:01")
End Sub



