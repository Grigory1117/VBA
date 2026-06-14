Attribute VB_Name = "ACC"

' ╔══════════════════════════════════════════════════════════════════╗
' ║                  主程序：一次登入爬取所有報表                        ║
' ╚══════════════════════════════════════════════════════════════════╝

Sub ACC()
    ' ========== 定義要爬取的報表列表 ==========
    ' 格式：ReportID, 目標SheetName, 來源SheetName
    Dim reportList As Variant
    reportList = Array( _
        Array("A1", "A1", "A1"), _
        Array("A2", "A2", "A2"), _
        Array("A3", "A3", "A3"), _
        Array("A4", "A4", "A4"), _
        Array("A5", "A5", "A5"), _
        Array("A6", "A6", "A6"), _
        Array("A7", "A7", "A7"), _
        Array("A8", "A8", "A8"), _
        Array("A9", "A9", "A9"), _
        Array("A0", "A0", "A0") _
    )
    ' ↑↑↑ 格式說明：
    ' 第1欄：ReportID (網址參數)
    ' 第2欄：目標Sheet名稱 (存到這個Excel的工作表名稱)
    ' 第3欄：來源Sheet名稱 (下載檔案中的工作表名稱)
    
    ' ========== 只輸入一次帳號密碼 ==========
    Dim username As String, password As String
    
    username = InputBox("請輸入帳號:", "登入")
    If username = "" Then Exit Sub

    password = InputBox("請輸入密碼:", "登入")
    If password = "" Then Exit Sub
    
    ' ========== 創建 IE  ==========
    Dim IE As Object
    Set IE = CreateObject("InternetExplorer.Application")
    IE.Visible = True
    
    ' ========== 登入（只需一次） ==========
    IE.navigate "https://example?"
    Do While IE.Busy Or IE.readyState <> 4: DoEvents: Loop

    IE.document.getElementById("uxUserId").Value = username
    IE.document.getElementById("uxPassword").Value = password
    IE.document.getElementById("uxSubmit").Click

    Do While IE.Busy Or IE.readyState <> 4: DoEvents: Loop
    Application.Wait Now + timeValue("00:00:03")

    ' ========== 迴圈爬取所有報表 ==========
    Dim i As Long
    Dim reportId As String
    Dim targetSheetName As String
    Dim sourceSheetName As String
    Dim successCount As Integer
    Dim failCount As Integer
    
    successCount = 0
    failCount = 0
    
    For i = LBound(reportList) To UBound(reportList)
        reportId = reportList(i)(0)
        targetSheetName = reportList(i)(1)
        sourceSheetName = reportList(i)(2)
        
        Debug.Print "==============================================="
        Debug.Print "========== 開始爬取 " & targetSheetName & " =========="
        Debug.Print "  ReportID: " & reportId
        Debug.Print "  目標Sheet: " & targetSheetName
        Debug.Print "  來源Sheet: " & sourceSheetName
        Debug.Print "==============================================="
        
        ' 呼叫通用爬蟲函數
        If CrawlReport(IE, reportId, targetSheetName, sourceSheetName) Then
            successCount = successCount + 1
            Debug.Print "========== " & targetSheetName & " 完成 =========="
        Else
            failCount = failCount + 1
            Debug.Print "========== " & targetSheetName & " 失敗 =========="
        End If
        
        ' 每個報表之間稍微等待
        Application.Wait Now + timeValue("00:00:02")
    Next i
    
    ' ========== 完成，關閉 IE ==========
    IE.Quit
    
    ' ========== 顯示結果 ==========
    Debug.Print "==============================================="
    Debug.Print "========== 全部完成 =========="
    Debug.Print "==============================================="
    Debug.Print "成功: " & successCount & " 個"
    Debug.Print "失敗: " & failCount & " 個"
    
    MsgBox "爬取完成！" & vbCrLf & vbCrLf & _
           "成功: " & successCount & " 個" & vbCrLf & _
           "失敗: " & failCount & " 個", vbInformation, "完成"
End Sub


' ╔══════════════════════════════════════════════════════════════════╗
' ║                      通用爬蟲函數（核心）                           ║
' ╚══════════════════════════════════════════════════════════════════╝

Function CrawlReport(IE As Object, reportId As String, targetSheetName As String, sourceSheetName As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' ========== 導航到報表頁面 ==========
    Dim reportUrl As String
    reportUrl = "https://example?reportId=" & reportId
    
    IE.navigate reportUrl
    Do While IE.Busy Or IE.readyState <> 4: DoEvents: Loop
    Application.Wait Now + timeValue("00:00:03")

    Debug.Print "========== 開始查詢 " & targetSheetName & " =========="

    ' ========== 點擊查詢按鈕 ==========
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

    ' ========== 取得表格資料 ==========
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

                    If serialCode <> "" And InStr(serialCode, ".") > 0 Then
                        Debug.Print "完整檔名: " & serialCode
                        
                        ' ========== 處理下載 ==========
                        Call ProcessReportWithDownload(serialCode, reportId, targetSheetName, sourceSheetName, IE, firstRow)
                        
                        Debug.Print targetSheetName & " 處理完成"
                        CrawlReport = True
                        Exit Function
                    Else
                        Debug.Print "檔案格式錯誤！取得: " & serialCode
                        CrawlReport = False
                        Exit Function
                    End If
                Else
                    Debug.Print "表格中沒有數據"
                    CrawlReport = False
                    Exit Function
                End If
            End If
        End If
    End If
    
    CrawlReport = False
    Exit Function
    
ErrorHandler:
    Debug.Print "發生錯誤: " & Err.Description
    CrawlReport = False
End Function


' ╔══════════════════════════════════════════════════════════════════╗
' ║                       資料提取與處理函數區                          ║
' ╚══════════════════════════════════════════════════════════════════╝

Function GetSerialCodeFromRow(row As Object) As String
    On Error Resume Next
    
    Dim Cells As Object
    Set Cells = row.getElementsByTagName("td")
    
    If Cells.Length >= 4 Then
        Dim dateCell As Object
        Dim timeCell As Object
        Dim fileTypeCell As Object
        Set dateCell = Cells(2)
        Set timeCell = Cells(3)
        Set fileTypeCell = Cells(4)
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
        Dim fileType As String

        If Len(cleanText) >= 30 Then
            dateValue = Left(cleanText, 8)
            timeValue = Mid(cleanText, 25, 6)
            fileType = Trim(fileTypeCell.innerText)

            Debug.Print "    提取日期: " & dateValue
            Debug.Print "    提取時間: " & timeValue
            Debug.Print "    檔案格式: " & fileType

            GetSerialCodeFromRow = dateValue & timeValue & "." & fileType
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


Sub ProcessReportWithDownload(serialNo As String, reportId As String, targetSheetName As String, sourceSheetName As String, IE As Object, row As Object)
    Debug.Print "========== 處理檔案 =========="
    Debug.Print "  ReportID: " & reportId
    Debug.Print "  流水碼: " & serialNo
    Debug.Print "  目標Sheet: " & targetSheetName
    Debug.Print "  來源Sheet: " & sourceSheetName
    
    Dim fileName As String
    Dim filePath As String
    Dim downloadFolder As String
    Dim currentUser As String
    Dim wbSource As Workbook
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    
    currentUser = Environ("USERNAME")
    
    ' ========== 檢查目標工作表是否存在，不存在則建立 ==========
    On Error Resume Next
    Set wsTarget = ThisWorkbook.Sheets(targetSheetName)
    If wsTarget Is Nothing Then
        Set wsTarget = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsTarget.Name = targetSheetName
        Debug.Print "  建立新工作表: " & targetSheetName
    End If
    On Error GoTo 0
    
    ' 設定下載路徑
    downloadFolder = "C:\Users\" & currentUser & "\Downloads\"
    fileName = reportId & "." & serialNo
    filePath = downloadFolder & fileName

    Debug.Print "  Windows 使用者: " & currentUser
    Debug.Print "  下載路徑: " & downloadFolder
    Debug.Print "  檔案路徑: " & filePath
    
    ' ========== 檢查檔案是否存在 ==========
    If Dir(filePath) = "" Then
        Debug.Print "  檔案不存在，開始自動下載..."

        Dim downloadSuccess As Boolean
        downloadSuccess = DownloadFileAutoClick(IE, row, downloadFolder, fileName)

        If Not downloadSuccess Then
            MsgBox "下載失敗：" & targetSheetName & vbCrLf & "請手動下載後重試。", vbExclamation
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
            MsgBox "開啟檔案失敗：" & targetSheetName & vbCrLf & Err.Description, vbCritical
            Exit Sub
        End If
        On Error GoTo 0

        ' ========== 根據指定的來源Sheet名稱開啟 ==========
        On Error Resume Next
        Set wsSource = wbSource.Sheets(sourceSheetName)
        
        If wsSource Is Nothing Then
            ' 如果找不到指定的Sheet，列出所有可用的Sheet
            Debug.Print "  ?? 找不到工作表: " & sourceSheetName
            Debug.Print "  可用的工作表："
            Dim ws As Worksheet
            For Each ws In wbSource.Worksheets
                Debug.Print "    - " & ws.Name
            Next ws
            
            ' 使用第一個工作表作為備案
            Set wsSource = wbSource.Sheets(1)
            Debug.Print "  ?? 改用第一個工作表: " & wsSource.Name
            
            MsgBox "注意：找不到工作表「" & sourceSheetName & "」" & vbCrLf & _
                   "已改用「" & wsSource.Name & "」", vbExclamation, "工作表名稱不符"
        Else
            Debug.Print "  ? 成功找到來源工作表: " & sourceSheetName
        End If
        On Error GoTo 0
        
        ' 清空目標工作表並複製資料
        wsTarget.Cells.Clear
        wsSource.Cells.Copy Destination:=wsTarget.Cells(1, 1)

        wbSource.Close SaveChanges:=False

        Debug.Print "  檔案處理完成"
    Else
        MsgBox "下載後仍找不到檔案：" & targetSheetName, vbExclamation
    End If
End Sub


' ========== 自動處理 IE 下載對話框 ==========
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
    
    ' 等待 IE 下載對話框出現
    Debug.Print "  等待下載對話框出現..."
    Application.Wait Now + timeValue("00:00:03")
    
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
        
        ' 方法3：直接按 Enter
        Debug.Print "    方法3: 直接 Enter"
        Application.SendKeys "{ENTER}", True
        Application.Wait Now + timeValue("00:00:01")
    End If
    
    ' 再檢查一次
    If Dir(downloadFolder & expectedFileName) = "" Then
        Debug.Print "    方法3 似乎也沒效果，嘗試方法4..."
        
        ' 方法4：Space 鍵
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


' ╔══════════════════════════════════════════════════════════════════╗
' ║                         通用輔助函數區                             ║
' ╚══════════════════════════════════════════════════════════════════╝

Sub WaitForIE(IE As Object)
    Do While IE.Busy Or IE.readyState <> 4
        DoEvents
    Loop
    Application.Wait Now + timeValue("00:00:01")
End Sub

