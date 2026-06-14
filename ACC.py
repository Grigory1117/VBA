
import time
import os
from pathlib import Path
from typing import List, Tuple
import getpass
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from webdriver_manager.microsoft import EdgeChromiumDriverManager
import logging
from datetime import datetime
from dateutil.relativedelta import relativedelta
import shutil
import win32com.client as win32

# 設定日誌
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(f'report_crawler_{datetime.now().strftime("%Y%m%d_%H%M%S")}.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)


class ReportCrawler:
    """報表爬蟲主類"""
    
    def __init__(self, template_path: str = "檢核.xlsx"):
        """初始化爬蟲"""
        self.driver = None
        self.download_folder = str(Path.home() / "Downloads")
        self.template_path = template_path
        self.target_workbook_path = None  # 稍後根據日期產生
        self.report_month = None  # 報表月份（YYYYMM）
        self.success_count = 0
        self.fail_count = 0
        self.excel_app = None  # Excel COM 物件
        
        # 檢查底稿是否存在
        if not os.path.exists(self.template_path):
            logger.error(f"找不到底稿檔案：{self.template_path}")
            raise FileNotFoundError(f"找不到底稿檔案：{self.template_path}")
    
    def setup_driver(self):
        """設定 Edge 瀏覽器驅動"""
        options = Options()
        prefs = {
            "download.default_directory": self.download_folder,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True
        }
        options.add_experimental_option("prefs", prefs)
        options.add_argument('--headless')
        service = Service(EdgeChromiumDriverManager().install())
        self.driver = webdriver.Edge(service=service, options=options)
        self.driver.maximize_window()
        logger.info("瀏覽器初始化完成")
    
    def setup_excel(self):
        """初始化 Excel COM 物件"""
        try:
            self.excel_app = win32.gencache.EnsureDispatch('Excel.Application')
            self.excel_app.Visible = False
            self.excel_app.DisplayAlerts = False
            logger.info("Excel COM 初始化完成")
        except Exception as e:
            logger.error(f"Excel COM 初始化失敗：{str(e)}")
            raise
    
    def calculate_report_month(self, date_value: str) -> str:
        """計算報表月份（含驗證）"""
        try:
            year = int(date_value[:4])
            month = int(date_value[4:6])
            
            # 驗證日期合理性
            if not (2020 <= year <= 2030):
                logger.warning(f"年份異常：{year}")
                # 使用當前月份作為備案
                return datetime.now().strftime("%Y%m")
            
            if not (1 <= month <= 12):
                logger.warning(f"月份異常：{month}")
                return datetime.now().strftime("%Y%m")
            
            current_date = datetime(year, month, 1)
            previous_month = current_date - relativedelta(months=1)
            
            report_month = previous_month.strftime("%Y%m")
            logger.info(f"計算報表月份：{date_value[:6]} - 1 = {report_month}")
            
            return report_month
        except Exception as e:
            logger.error(f"計算報表月份失敗：{str(e)}")
            # 備用方案：使用當前月份
            return datetime.now().strftime("%Y%m")
    
    def create_target_workbook(self, report_month: str):
        """
        從底稿建立目標檔案
        
        Args:
            report_month: 報表月份（YYYYMM）
        """
        # 產生目標檔案名稱
        template_dir = os.path.dirname(os.path.abspath(self.template_path)) or "."
        template_name = os.path.splitext(os.path.basename(self.template_path))[0]
        self.target_workbook_path = os.path.join(template_dir, f"{template_name}{report_month}.xlsx")
        
        # 如果目標檔案已存在，先刪除
        if os.path.exists(self.target_workbook_path):
            logger.info(f"目標檔案已存在，將覆蓋：{self.target_workbook_path}")
            try:
                os.remove(self.target_workbook_path)
            except:
                pass
        
        # 複製底稿
        shutil.copy2(self.template_path, self.target_workbook_path)
        logger.info(f"已建立目標檔案：{os.path.basename(self.target_workbook_path)}")
    
    def login(self, username: str, password: str, login_url: str = "https://example?"):
        """登入網站"""
        try:
            logger.info("開始登入...")
            self.driver.get(login_url)
            wait = WebDriverWait(self.driver, 10)
            
            username_input = wait.until(EC.presence_of_element_located((By.ID, "uxUserId")))
            password_input = self.driver.find_element(By.ID, "uxPassword")
            submit_button = self.driver.find_element(By.ID, "uxSubmit")
            
            username_input.clear()
            username_input.send_keys(username)
            password_input.clear()
            password_input.send_keys(password)
            submit_button.click()
            
            time.sleep(3)
            logger.info("登入成功")
            return True
            
        except Exception as e:
            logger.error(f"登入失敗：{str(e)}")
            return False
    
    def crawl_report(self, report_id: str, target_sheet_name: str, source_sheet_name: str) -> bool:
        """爬取單個報表"""
        try:
            logger.info("=" * 50)
            logger.info(f"開始爬取：{target_sheet_name} (ReportID: {report_id})")
            
            # 導航到報表頁面
            report_url = f"https://example?reportId={report_id}"
            self.driver.get(report_url)
            time.sleep(3)
            
            # 點擊查詢按鈕
            wait = WebDriverWait(self.driver, 10)
            query_button = wait.until(EC.element_to_be_clickable((By.ID, "queryButton")))
            query_button.click()
            time.sleep(5)
            
            # 取得表格資料
            panels = self.driver.find_elements(By.CLASS_NAME, "col-sm-12")
            if not panels:
                logger.warning("找不到資料面板")
                return False
            
            tables = panels[0].find_elements(By.TAG_NAME, "table")
            if not tables:
                logger.warning("找不到資料表格")
                return False
            
            tbody = tables[0].find_elements(By.TAG_NAME, "tbody")
            if not tbody:
                logger.warning("找不到表格內容")
                return False
            
            rows = tbody[0].find_elements(By.TAG_NAME, "tr")
            if not rows:
                logger.warning("表格中沒有數據")
                return False
            
            # 從第一行取得流水碼
            first_row = rows[0]
            serial_code, date_value = self.get_serial_code_from_row(first_row)
            
            if not serial_code or '.' not in serial_code:
                logger.error(f"無法取得流水碼：{serial_code}")
                return False
            
            logger.info(f"取得流水碼：{serial_code}")
            
            # 如果是第一個報表，計算報表月份並建立目標檔案
            if self.report_month is None and date_value:
                self.report_month = self.calculate_report_month(date_value)
                self.create_target_workbook(self.report_month)
            
            # 處理下載
            success = self.process_report_with_download(
                serial_code, report_id, target_sheet_name, source_sheet_name, first_row
            )
            
            if success:
                logger.info(f"✓ {target_sheet_name} 處理完成")
            return success
            
        except Exception as e:
            logger.error(f"爬取失敗：{str(e)}")
            return False
    
    def get_serial_code_from_row(self, row) -> Tuple[str, str]:
        """
        從表格行中提取流水碼
        
        Returns:
            Tuple[str, str]: (流水碼, 日期值)
        """
        try:
            cells = row.find_elements(By.TAG_NAME, "td")
            
            if len(cells) < 5:
                logger.warning("儲存格數量不足")
                return "", ""
            
            # 按照原始 VBA 邏輯：cells[2]=日期, cells[3]=時間, cells[4]=副檔名
            date_cell = cells[2]
            time_cell = cells[3]
            file_type_cell = cells[4]
            
            # 合併日期時間文字，只保留數字
            raw_text = date_cell.text.strip() + time_cell.text.strip()
            clean_text = ''.join(filter(str.isdigit, raw_text))
            
            if len(clean_text) >= 14:
                date_value = clean_text[:8]      # 前8碼：日期
                time_value = clean_text[8:14]    # 第8-14碼：時間
                file_type = file_type_cell.text.strip().upper()  # 統一大寫
                
                serial_code = f"{date_value}{time_value}.{file_type}"
                return serial_code, date_value
            else:
                logger.warning(f"日期時間長度不足：{len(clean_text)} (需要14)")
                return "", ""
                
        except Exception as e:
            logger.error(f"提取流水碼錯誤：{str(e)}")
            return "", ""
    
    def process_report_with_download(self, serial_no: str, report_id: str, 
                                     target_sheet_name: str, source_sheet_name: str, row) -> bool:
        """處理報表下載和資料複製（使用 COM 保留格式）"""
        try:
            # 組合檔案名稱：ReportID.流水碼
            file_name = f"{report_id}.{serial_no}"
            file_path = os.path.join(self.download_folder, file_name)
            
            logger.info(f"處理檔案：{file_name}")
            
            # 檢查檔案是否存在（忽略大小寫）
            actual_file_path = self.find_file_case_insensitive(self.download_folder, file_name)
            
            if not actual_file_path:
                logger.info("開始下載...")
                download_success = self.download_file(row, file_path)
                
                if not download_success:
                    logger.error("下載失敗")
                    return False
                
                logger.info("下載完成")
                # 重新檢查檔案
                actual_file_path = self.find_file_case_insensitive(self.download_folder, file_name)
                
                if not actual_file_path:
                    logger.error("下載後找不到檔案")
                    return False
            else:
                logger.info("檔案已存在，跳過下載")
            
            # 判斷檔案類型
            file_extension = os.path.splitext(actual_file_path)[1].lower()
            
            # 如果是 .xls 格式，先轉換成 .xlsx（使用 COM 保留格式）
            if file_extension == '.xls':
                logger.info("偵測到 XLS 格式，轉換為 XLSX（保留格式）...")
                xlsx_path = actual_file_path + 'x'  # .xls -> .xlsx
                
                if not self.convert_xls_to_xlsx_with_format(actual_file_path, xlsx_path):
                    logger.error("XLS 轉換失敗")
                    return False
                
                actual_file_path = xlsx_path
                logger.info("轉換完成")
            
            # 使用 COM 複製工作表（保留完整格式）
            logger.info(f"來源工作表：{source_sheet_name} → 目標工作表：{target_sheet_name}")
            
            success = self.copy_worksheet_with_format(
                actual_file_path, 
                source_sheet_name, 
                self.target_workbook_path, 
                target_sheet_name
            )
            
            if success:
                logger.info("資料複製完成（格式已保留）")
            
            return success
            
        except Exception as e:
            logger.error(f"處理檔案失敗：{str(e)}")
            import traceback
            logger.error(traceback.format_exc())
            return False
    
    def convert_xls_to_xlsx_with_format(self, xls_path: str, xlsx_path: str) -> bool:
        """使用 Excel COM 將 XLS 轉換為 XLSX（保留格式）"""
        wb = None
        try:
            # 開啟 XLS 檔案
            wb = self.excel_app.Workbooks.Open(os.path.abspath(xls_path))
            
            # 另存為 XLSX 格式 (51 = xlOpenXMLWorkbook)
            wb.SaveAs(os.path.abspath(xlsx_path), FileFormat=51)
            
            logger.info(f"已轉換（格式保留）：{os.path.basename(xls_path)} → {os.path.basename(xlsx_path)}")
            return True
            
        except Exception as e:
            logger.error(f"XLS 轉換錯誤：{str(e)}")
            return False
        finally:
            if wb:
                wb.Close(SaveChanges=False)
    
    def copy_worksheet_with_format(self, source_file: str, source_sheet: str, 
                                target_file: str, target_sheet: str) -> bool:
        """使用 Excel COM 複製工作表內容（保留格式，不破壞公式參照）"""
        source_wb = None
        target_wb = None
        try:
            # 開啟來源和目標檔案
            source_wb = self.excel_app.Workbooks.Open(os.path.abspath(source_file))
            target_wb = self.excel_app.Workbooks.Open(os.path.abspath(target_file))
            
            # 找到來源工作表
            source_ws = None
            for ws in source_wb.Worksheets:
                if ws.Name == source_sheet:
                    source_ws = ws
                    break
            
            if not source_ws:
                logger.warning(f"找不到工作表 [{source_sheet}]，改用第一個工作表")
                source_ws = source_wb.Worksheets(1)
                logger.warning(f"使用工作表：{source_ws.Name}")
            
            # 找到或建立目標工作表
            target_ws = None
            for ws in target_wb.Worksheets:
                if ws.Name == target_sheet:
                    target_ws = ws
                    break
            
            if not target_ws:
                # 如果目標工作表不存在，建立新的
                target_ws = target_wb.Worksheets.Add()
                target_ws.Name = target_sheet
                logger.info(f"建立新工作表：{target_sheet}")
            else:
                # 如果存在，清空內容（但保留工作表，避免破壞參照）
                target_ws.Cells.Clear()
                logger.info(f"清空工作表內容：{target_sheet}")
            
            # 複製整個已使用範圍（包含格式）
            if source_ws.UsedRange.Count > 0:
                source_ws.UsedRange.Copy()
                target_ws.Range("A1").PasteSpecial(Paste=-4163)  # xlPasteAll = -4163
                self.excel_app.CutCopyMode = False
                
                # 複製列高和欄寬
                for i in range(1, source_ws.UsedRange.Rows.Count + 1):
                    target_ws.Rows(i).RowHeight = source_ws.Rows(i).RowHeight
                
                for i in range(1, source_ws.UsedRange.Columns.Count + 1):
                    target_ws.Columns(i).ColumnWidth = source_ws.Columns(i).ColumnWidth
            
            # 儲存目標檔案
            target_wb.Save()
            
            return True
            
        except Exception as e:
            logger.error(f"複製工作表錯誤：{str(e)}")
            import traceback
            logger.error(traceback.format_exc())
            return False
        finally:
            if source_wb:
                source_wb.Close(SaveChanges=False)
            if target_wb:
                target_wb.Close(SaveChanges=False)
    
    def find_file_case_insensitive(self, folder: str, filename: str) -> str:
        """不區分大小寫地尋找檔案"""
        try:
            filename_lower = filename.lower()
            for file in os.listdir(folder):
                if file.lower() == filename_lower:
                    return os.path.join(folder, file)
            return ""
        except Exception:
            return ""
    
    def download_file(self, row, file_path: str, timeout: int = 60) -> bool:
        """自動下載檔案"""
        try:
            # 找到下載按鈕
            cells = row.find_elements(By.TAG_NAME, "td")
            download_cell = cells[-2]
            buttons = download_cell.find_elements(By.TAG_NAME, "button")
            
            # 尋找下載按鈕
            download_btn = None
            for button in buttons:
                if "下載" in button.text or button.is_displayed():
                    download_btn = button
                    break
            
            if not download_btn:
                logger.error("找不到下載按鈕")
                return False
            
            # 點擊下載
            download_btn.click()
            logger.info("已點擊下載按鈕")
            
            # 等待下載完成
            wait_count = 0
            while wait_count < timeout:
                time.sleep(1)
                wait_count += 1
                
                # 檢查檔案（忽略大小寫）
                actual_path = self.find_file_case_insensitive(
                    os.path.dirname(file_path), 
                    os.path.basename(file_path)
                )
                
                if actual_path:
                    # 檢查檔案大小是否穩定
                    file_size1 = os.path.getsize(actual_path)
                    time.sleep(2)
                    file_size2 = os.path.getsize(actual_path)
                    
                    if file_size1 == file_size2 and file_size1 > 0:
                        logger.info(f"下載完成 (等待 {wait_count} 秒，大小：{file_size1:,} bytes)")
                        return True
                
                if wait_count % 10 == 0:
                    logger.info(f"等待下載... ({wait_count} 秒)")
            
            logger.error(f"下載逾時 (超過 {timeout} 秒)")
            return False
            
        except Exception as e:
            logger.error(f"下載錯誤：{str(e)}")
            return False
    
    def open_target_file(self):
        """開啟目標檔案"""
        try:
            if self.target_workbook_path and os.path.exists(self.target_workbook_path):
                logger.info(f"開啟檔案：{os.path.basename(self.target_workbook_path)}")
                os.startfile(os.path.abspath(self.target_workbook_path))
            else:
                logger.warning("目標檔案不存在")
        except Exception as e:
            logger.error(f"開啟檔案失敗：{str(e)}")
    
    def run(self, report_list: List[Tuple[str, str, str]], username: str = None, password: str = None):
        """執行完整的爬取流程"""
        try:
            # 取得帳號密碼
            if not username:
                username = input("請輸入帳號: ")
            if not password:
                password = getpass.getpass("請輸入密碼: ")
            
            # 初始化 Excel COM
            self.setup_excel()
            
            # 設定瀏覽器並登入
            self.setup_driver()
            
            if not self.login(username, password):
                logger.error("登入失敗，程式終止")
                return
            
            # 爬取所有報表
            logger.info(f"開始爬取 {len(report_list)} 個報表")
            
            for report_id, target_sheet, source_sheet in report_list:
                if self.crawl_report(report_id, target_sheet, source_sheet):
                    self.success_count += 1
                else:
                    self.fail_count += 1
                
                time.sleep(2)  # 每個報表間隔
            
            # 顯示結果
            logger.info("=" * 50)
            logger.info(f"爬取完成！成功：{self.success_count} 個，失敗：{self.fail_count} 個")
            logger.info("=" * 50)
            
            print(f"\n✓ 爬取完成！")
            print(f"  成功：{self.success_count} 個")
            print(f"  失敗：{self.fail_count} 個")
            
            if self.target_workbook_path:
                print(f"  檔案：{os.path.basename(self.target_workbook_path)}")
            
            # 開啟目標檔案
            if self.success_count > 0:
                self.open_target_file()
            
        except KeyboardInterrupt:
            logger.info("使用者中斷程式")
        except Exception as e:
            logger.error(f"執行錯誤：{str(e)}")
            import traceback
            logger.error(traceback.format_exc())
        finally:
            # 關閉 Excel
            # if self.excel_app:
            #     self.excel_app.Quit()
            #     logger.info("Excel 已關閉")
            
            # 關閉瀏覽器
            if self.driver:
                self.driver.quit()
                logger.info("瀏覽器已關閉")


def main():
    """主程式入口"""
    
    # 定義要爬取的報表列表
    # 格式：(ReportID, 目標SheetName, 來源SheetName)
    report_list = [
        ("A1", "A1", "A1"),
        ("A2", "A2", "A2"),
        ("A3", "A3", "A3"),
        ("A4", "A4", "A4"),
        ("A5", "A5", "A5"),
        ("A6", "A6", "A6"),
        ("A7", "A7", "A7"),
        ("A8", "A8", "A8"),
        ("A9", "A9", "A9"),
        ("A0", "A0", "A0"),
    ]
    
    # 建立爬蟲實例並執行
    crawler = ReportCrawler(template_path="檢核.xlsx")
    crawler.run(report_list)


if __name__ == "__main__":
    main()