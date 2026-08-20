# -*- coding: utf-8 -*-
"""
📄 電子公文歸檔自動化系統 (Official Document Auto-Archiver) v2.0
專為公務與教育機構行政同仁設計之自動化歸檔輔助工具。
透過 Python、Selenium 與 Tkinter 實現批次選擇與全自動點擊歸檔。
"""

import os
import sys
import time
import subprocess
import threading
import traceback
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext

# Selenium 相關模組
try:
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.support.ui import WebDriverWait, Select
    from selenium.webdriver.support import expected_conditions as EC
except ImportError:
    webdriver = None

# ==========================================
# 階段零：版本資訊與更新檢查
# ==========================================
CURRENT_VERSION = "v1.3.0"
GITHUB_REPO = "ChenYuChunEric/Taipei-Official-Doc-Auto-Archive-"

def check_for_updates(quiet=False):
    """檢查 GitHub Releases 是否有新版本"""
    try:
        import urllib.request
        import json
        url = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req, timeout=5) as response:
            if response.status == 200:
                data = json.loads(response.read().decode('utf-8'))
                latest_version = data.get('tag_name', '')
                if latest_version and latest_version != CURRENT_VERSION:
                    html_url = data.get('html_url', f"https://github.com/{GITHUB_REPO}/releases")
                    return latest_version, html_url
    except Exception as e:
        print(f"檢查更新失敗: {e}")
    return None, None

# ==========================================
# 階段一：分類號大辭典 (Master Categories)
# ==========================================
MASTER_CATEGORIES = {
    "03010101": "03010101-綜合業務(3年)", "03010102": "03010102-防護工作(3年)", "03010104": "03010104-會議及史料(永久)",
    "03010201": "03010201-一般文書管理(3年)", "03010202": "03010202-用印信申請(3年)", "03010203": "03010203-公文收發(30年)",
    "03010301": "03010301-一般檔案管理(3年)", "03010302": "03010302-檔案借調(5年)", "03010303": "03010303-檔案清理及移交(永久)",
    "03010401": "03010401-採購管理(5年)", "03010402": "03010402-採購爭議(20年)", "03010403": "03010403-營繕工程(永久)",
    "030105": "030105-出納管理目(10年)", "030106": "030106-財產管理目", "03010601": "03010601-動產管理(10年)",
    "03010602": "03010602-土地管理(永久)", "03010701": "03010701-物品管理(3年)", "03010702": "03010702-廢品管理(3年)",
    "030108": "030108-車輛管理目(3年)", "030109": "030109-廳舍管理目", "03010901": "03010901-公共與消防安全(3年)",
    "03010902": "03010902-災害與破壞事件(15年)", "03010903": "03010903-宿舍管理(永久)", "03011001": "03011001-法令宣導(3年)",
    "03011002": "03011002-法令及釋疑(10年)", "03020101": "03020101-綜合業務(3年)", "03020201": "03020201-首長交接及人力評鑑(20年)",
    "03020202": "03020202-職務歸系(50年)", "03020203": "03020203-組織編制(永久)", "03020301": "03020301-派免遷調核薪(50年)",
    "03020302": "03020302-甄選聘僱(50年)", "03020401": "03020401-一般獎懲(10年)", "03020402": "03020402-考績（核）、重大獎懲、停職(50年)",
    "03020403": "03020403-獎章(50年)", "03020404": "03020404-績優人員(5年)", "030205": "030205-訓練進修考察目(10年)",
    "030206": "030206-差勤管理目(5年)", "030207": "030207-保障目(25年)", "030208": "030208-俸給待遇目(10年)",
    "03020901": "03020901-福利、津貼、給與(10年)", "03020902": "03020902-輔購（建）住宅(30年)", "030210": "030210-保險目(10年)",
    "03021101": "03021101-退休照護(5年)", "03021103": "03021103-資遣及退撫基金(50年)", "03021104": "03021104-一次退休(50年)",
    "03021105": "03021105-退休撫卹(永久)", "030212": "030212-人事資料、服務目(10年)", "03021301": "03021301-法令宣導(3年)",
    "03021302": "03021302-法令及釋疑(10年)", "03030101": "03030101-綜合業務(3年)", "030302": "030302-預算目(10年)",
    "030303": "030303-決算目(10年)", "03030501": "03030501-會計相關規定(10年)", "03030502": "03030502-會計報告、簿籍(15年)",
    "03030601": "03030601-法規宣導(3年)", "03030602": "03030602-法規及釋疑(10年)", "03750101": "03750101-綜合業務(3年)",
    "03750201": "03750201-教務會議、教師研習(5年)", "03750202": "03750202-教學活動(5年)", "03750203": "03750203-課務處理(5年)",
    "03750204": "03750204-課後留園(5年)", "03750205": "03750205-教學研究及視導(10年)", "03750301": "03750301-證明文件核發(1年)",
    "03750302": "03750302-教育統計(5年)", "03750303": "03750303-獎助學金及就學優待(5年)", "03750304": "03750304-招生宣導及入學(10年)",
    "03750305": "03750305-成績管理(30年)", "03750306": "03750306-學籍管理(永久)", "03750401": "03750401-資訊研習及競賽(3年)",
    "03750402": "03750402-資通安全(3年)", "03750403": "03750403-資訊教學(5年)", "03750404": "03750404-資訊設備管理及維護(10年)",
    "03750501": "03750501-科學競賽(5年)", "03750502": "03750502-設備管理及維護(10年)", "03750503": "03750503-教材選用(20年)",
    "03750601": "03750601-活動競賽及藝術才能(5年)", "03750701": "03750701-法規宣導(3年)", "03750702": "03750702-法規及釋疑(10年)",
    "03760101": "03760101-綜合業務(3年)", "03760201": "03760201-會議 、導師制度(5年)", "03760202": "03760202-學生校內外活動(5年)",
    "03760203": "03760203-學生獎懲申訴及救濟(20年)", "03760301": "03760301-交通安全(3年)", "03760302": "03760302-校園安全(5年)",
    "03760303": "03760303-學生事務方案推展(5年)", "03760304": "03760304-生活教育輔導(10年)", "037604": "037604-學校體育目(5年)",
    "03760501": "03760501-環境衛生、營養補助(3年)", "03760502": "03760502-衛生保健、餐飲管理(10年)", "03760503": "03760503-學生平安保險(20年)",
    "03760601": "03760601-法規宣導(3年)", "03760602": "03760602-法規及釋疑(10年)", "03770101": "03770101-綜合業務(3年)",
    "03770201": "03770201-學習輔導及紀錄(3年)", "03770202": "03770202-輔導會議及活動(5年)", "03770203": "03770203-中輟生輔導及紀錄(10年)",
    "03770204": "03770204-個案輔導(20年)", "03770301": "03770301-特殊教育教學及活動(5年)", "03770302": "03770302-特教資源(10年)",
    "03770303": "03770303-特教工作鑑定與安置(10年)", "03770401": "03770401-輔導圖書管理(10年)", "03770402": "03770402-資料管理(10年)",
    "03770501": "03770501-技藝競賽與檢定(3年)", "03770601": "03770601-成人及推廣教育(3年)", "03770602": "03770602-親職教育(3年)",
    "03770603": "03770603-社區關係(5年)", "03780101": "03780101-綜合業務(3年)", "037803": "037803-法令規章目",
    "03790101": "03790101-綜合業務(3年)", "037905": "037905-家長會目(10年)", "03800101": "03800101-綜合業務(3年)",
    "03800201": "03800201-技術、讀者服務(3年)", "03800202": "03800202-圖書推薦(3年)", "03800203": "03800203-館際活動(3年)",
    "03800301": "03800301-法令宣導(3年)", "03810101": "03810101-綜合業務(3年)"
}

class CategoryManager:
    """管理常用分類項目與 categories.txt 讀寫"""
    
    def __init__(self, base_path=None):
        if base_path is None:
            if getattr(sys, 'frozen', False):
                base_path = os.path.dirname(sys.executable)
            else:
                base_path = os.path.dirname(os.path.abspath(__file__))
        self.config_file = os.path.join(base_path, 'categories.txt')
        self.category_map = {}
        self.load_categories()

    def load_categories(self):
        """讀取 categories.txt，若不存在則自動生成"""
        default_codes = [
            "03750401", "03750402", "03750403", "03750404", 
            "03750501", "03750502", "03010203"
        ]
        self.category_map.clear()

        if not os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'w', encoding='utf-8') as f:
                    f.write("# ==========================================\n")
                    f.write("# 電子公文小幫手 - 常用分類設定檔\n")
                    f.write("# 請在下方輸入您常使用的分類代碼數字（每行一個）\n")
                    f.write("# 程式啟動時會自動為您轉換為完整的分類名稱與案號！\n")
                    f.write("# ==========================================\n\n")
                    for code in default_codes:
                        f.write(f"{code}\n")
                
                for code in default_codes:
                    display_name = MASTER_CATEGORIES.get(code, f"{code} (分類號)")
                    self.category_map[display_name] = code
            except Exception as e:
                print(f"創建預設 categories.txt 失敗: {e}")
        else:
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    for line in f:
                        code = line.strip()
                        if not code or code.startswith("#"):
                            continue
                        display_name = MASTER_CATEGORIES.get(code, f"{code} (分類號)")
                        self.category_map[display_name] = code
            except Exception as e:
                print(f"讀取 categories.txt 失敗: {e}")

        # 加入不安裝歸檔的輔助選項
        self.category_map["手動處理 (不自動歸檔)"] = None
        return self.category_map

# ==========================================
# 階段二：Selenium 自動化核心 Engine
# ==========================================
class SeleniumEngine:
    """負責操控 Chrome Debugger 與電子公文系統介面"""
    
    def __init__(self, port=9222):
        self.port = port
        self.driver = None

    def launch_chrome_process(self):
        """開啟獨立 Chrome 分身 (啟用 Debug Port 9222)"""
        chrome_paths = [
            r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
        ]
        
        chrome_path = None
        for path in chrome_paths:
            if os.path.exists(path):
                chrome_path = path
                break
                
        if not chrome_path:
            return False, "找不到 Google Chrome 瀏覽器，請確認電腦是否已安裝 Chrome。"

        user_home = os.path.expanduser("~")
        user_data_dir = os.path.join(user_home, "SeleniumChromeProfile")

        try:
            subprocess.Popen([
                chrome_path, 
                "https://edoc.gov.taipei/tcqb/index.jsp#",
                f"--remote-debugging-port={self.port}", 
                f"--user-data-dir={user_data_dir}"
            ])
            return True, "成功啟動專用 Chrome 瀏覽器！"
        except Exception as e:
            return False, f"無法啟動 Chrome: {e}"

    def connect_driver(self):
        """附加到已啟動的 Debugger Chrome"""
        if self.driver:
            try:
                self.driver.current_url
                return True
            except:
                self.driver = None

        try:
            chrome_options = Options()
            chrome_options.add_experimental_option("debuggerAddress", f"127.0.0.1:{self.port}")
            self.driver = webdriver.Chrome(options=chrome_options)
            return True
        except Exception as e:
            return False

    def fetch_pending_docs(self, log_func=print):
        """讀取待辦公文清單 (支援多頁自動抓取)"""
        if not self.connect_driver():
            raise Exception("無法連線至 Chrome (Port 9222)。請確認專用 Chrome 視窗已啟動。")

        log_func("🔍 正在定位電子公文系統主畫面與待辦清單 Frame...")
        self.driver.switch_to.default_content()

        # 等待切換至 dTreeContent Frame
        try:
            WebDriverWait(self.driver, 5).until(
                EC.frame_to_be_available_and_switch_to_it((By.ID, "dTreeContent"))
            )
        except Exception:
            # 嘗試尋找其他可能名稱的 iframe
            iframes = self.driver.find_elements(By.TAG_NAME, "iframe")
            switched = False
            for iframe in iframes:
                frame_id = iframe.get_attribute("id") or iframe.get_attribute("name") or ""
                if "content" in frame_id.lower() or "tree" in frame_id.lower() or "main" in frame_id.lower():
                    self.driver.switch_to.frame(iframe)
                    switched = True
                    break
            if not switched:
                log_func("⚠️ 未能自動找到 dTreeContent Frame，嘗試在主文件讀取...")

        pending_docs = []
        page_count = 1
        max_pages = 20

        while page_count <= max_pages:
            log_func(f"📄 正在抓取第 {page_count} 頁待辦公文...")
            rows = self.driver.find_elements(By.XPATH, "//tbody[@id='listTBODY']/tr")
            if not rows:
                # 備用 table tr 掃描
                rows = self.driver.find_elements(By.XPATH, "//table[contains(@class, 'table')]//tr[td//input[@name='ids']]")

            current_page_count = 0
            for row in rows:
                try:
                    checkbox = row.find_element(By.NAME, "ids")
                    doc_id = checkbox.get_attribute("value")
                    if not doc_id:
                        continue

                    title_text = "無主旨"
                    try:
                        title_el = row.find_element(By.XPATH, ".//td[@data-th='主旨' or @data-th='主旨摘要']//span[@id='mainSpan']")
                        title_text = title_el.text.strip()
                    except:
                        try:
                            title_el = row.find_element(By.XPATH, ".//span[@id='mainSpan']")
                            title_text = title_el.text.strip()
                        except:
                            tds = row.find_elements(By.TAG_NAME, "td")
                            if len(tds) > 2:
                                title_text = tds[2].text.strip()

                    pending_docs.append({"id": doc_id, "title": title_text})
                    current_page_count += 1
                except Exception:
                    continue

            log_func(f"  └ 本頁成功抓取 {current_page_count} 筆。")

            # 嘗試翻到下一頁
            next_btns = self.driver.find_elements(
                By.XPATH, 
                "//a[contains(text(), '下一頁') or contains(@title, '下一頁') or contains(text(), '>')] | "
                "//input[contains(@value, '下一頁')] | //button[contains(text(), '下一頁')]"
            )
            clicked_next = False
            for btn in next_btns:
                if btn.is_displayed() and btn.is_enabled() and "disabled" not in (btn.get_attribute("class") or ""):
                    try:
                        self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn)
                        time.sleep(0.3)
                        btn.click()
                    except:
                        self.driver.execute_script("arguments[0].click();", btn)
                    time.sleep(2)
                    clicked_next = True
                    page_count += 1
                    break
            
            if not clicked_next:
                break

        return pending_docs

    def select_doc_checkbox(self, target_id, log_func=print):
        """在頁面上精確找到指定 ID 的公文並勾選 (使用 iCheck JS)"""
        max_pages = 20
        for page in range(max_pages):
            checkboxes = self.driver.find_elements(By.NAME, "ids")
            for cb in checkboxes:
                if cb.get_attribute("value") == target_id:
                    smart_check_script = f"""
                        var targetId = '{target_id}';
                        var checkboxes = document.getElementsByName('ids');
                        for(var i=0; i<checkboxes.length; i++) {{
                            var cb = checkboxes[i];
                            if(cb.value === targetId) {{
                                if(!cb.checked) {{ 
                                    if(typeof $ !== 'undefined' && $.fn.iCheck) {{ $(cb).iCheck('check'); }}
                                    cb.checked = true; 
                                }}
                            }} else {{
                                if(cb.checked) {{
                                    if(typeof $ !== 'undefined' && $.fn.iCheck) {{ $(cb).iCheck('uncheck'); }}
                                    cb.checked = false; 
                                }}
                            }}
                        }}
                    """
                    self.driver.execute_script(smart_check_script)
                    return True
                    
            # 翻頁尋找
            next_btns = self.driver.find_elements(
                By.XPATH, 
                "//a[contains(text(), '下一頁') or contains(@title, '下一頁') or contains(text(), '>')] | "
                "//input[contains(@value, '下一頁')] | //button[contains(text(), '下一頁')]"
            )
            clicked_next = False
            for btn in next_btns:
                if btn.is_displayed() and btn.is_enabled() and "disabled" not in (btn.get_attribute("class") or ""):
                    try:
                        self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn)
                        time.sleep(0.3)
                        btn.click()
                    except:
                        self.driver.execute_script("arguments[0].click();", btn)
                    time.sleep(2)
                    clicked_next = True
                    break
            if not clicked_next:
                break
                
        return False

    def process_single_doc_archive(self, doc_id, doc_title, category_code, log_func=print):
        """
        處理單筆公文自動歸檔完整流程：
        1. 勾選公文
        2. 點擊「存查」
        3. 選擇/填寫分類號
        4. 點擊「附件歸檔」按鈕
        5. 等待並自動關閉「附件歸檔」彈出視窗
        6. 關閉附件視窗後點擊「確定存檔」
        7. 監測與等待憑證簽章或結果
        """
        log_func(f"\n🚀 [開始處理] 公文 ID: {doc_id}")
        log_func(f"   主旨: {doc_title}")
        log_func(f"   目標分類號: {category_code}")

        if not self.connect_driver():
            raise Exception("與 Chrome 連線中斷")

        # 1. 確保切換至包含列表的 Frame
        self.driver.switch_to.default_content()
        try:
            WebDriverWait(self.driver, 5).until(
                EC.frame_to_be_available_and_switch_to_it((By.ID, "dTreeContent"))
            )
        except Exception:
            pass

        # 2. 勾選指定公文
        if not self.select_doc_checkbox(doc_id, log_func):
            log_func(f"❌ 錯誤: 頁面上找不到公文 ID {doc_id}，嘗試繼續下一筆。")
            return False

        log_func("  ✓ 已成功勾選該筆公文。")
        time.sleep(0.5)

        # 3. 點擊「存查」按鈕
        save_btns = self.driver.find_elements(
            By.XPATH, 
            "//a[contains(text(), '存查') or contains(@value, '存查')] | "
            "//input[contains(@value, '存查')] | //button[contains(text(), '存查')]"
        )
        clicked_save = False
        for btn in save_btns:
            if btn.is_displayed() and btn.is_enabled():
                try:
                    btn.click()
                except:
                    self.driver.execute_script("arguments[0].click();", btn)
                clicked_save = True
                break

        if not clicked_save:
            log_func("⚠️ 找不到『存查』按鈕，嘗試觸發頁面上的 default 存查 function...")
            try:
                self.driver.execute_script("if(typeof doSave !== 'undefined'){ doSave(); } else if(typeof saveDoc !== 'undefined'){ saveDoc(); }")
                clicked_save = True
            except Exception as e:
                log_func(f"❌ 點擊『存查』失敗: {e}")
                return False

        log_func("  ✓ 已點擊『存查』按鈕，等待彈出歸檔/設定視窗...")
        time.sleep(2)

        # 4. 【使用者精確順序】點擊「附件歸檔」按鈕並處理開啟的新分頁
        log_func("📌 尋找並點擊『附件歸檔』按鈕...")
        original_window = self.driver.current_window_handle
        old_windows = set(self.driver.window_handles)

        attach_archive_clicked = False
        attach_btns = self.driver.find_elements(
            By.XPATH, 
            "//a[contains(text(), '附件歸檔')] | //button[contains(text(), '附件歸檔')] | "
            "//input[contains(@value, '附件歸檔')] | //*[contains(text(), '附件歸檔') and (self::a or self::button or self::span)]"
        )
        for btn in attach_btns:
            if btn.is_displayed() and btn.is_enabled():
                try:
                    self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn)
                    time.sleep(0.3)
                    btn.click()
                except:
                    self.driver.execute_script("arguments[0].click();", btn)
                attach_archive_clicked = True
                break

        if attach_archive_clicked:
            log_func("  ✓ 已成功點擊『附件歸檔』按鈕！等待開啟新分頁...")
            time.sleep(2)

            # 切換至開啟的「附件歸檔」新分頁，並關閉該新分頁
            log_func("📌 偵測並處理跳出的『附件歸檔』新分頁...")
            new_windows = set(self.driver.window_handles) - old_windows

            if new_windows:
                new_tab = list(new_windows)[0]
                self.driver.switch_to.window(new_tab)
                log_func("  ✓ 已切換至『附件歸檔』新分頁，正在自動完成關閉...")
                time.sleep(1.5)
                try:
                    self.driver.close()
                    log_func("  ✓ 已成功關閉『附件歸檔』新分頁！")
                except Exception as e:
                    log_func(f"  ⚠️ 關閉新分頁提示: {e}")
                
                # 切回原本的主公文視窗
                self.driver.switch_to.window(original_window)
                log_func("  ✓ 已切回到主公文畫面。")
            else:
                log_func("  ℹ️ 未偵測到獨立新分頁，嘗試檢查彈跳視窗與對話框關閉...")
                try:
                    alert = self.driver.switch_to.alert
                    alert.accept()
                    log_func("  ✓ 已自動確認提示彈窗。")
                except:
                    pass

                close_btns = self.driver.find_elements(
                    By.XPATH, 
                    "//button[contains(text(), '關閉') or contains(text(), '確定') or contains(@class, 'close')] | "
                    "//a[contains(text(), '關閉') or contains(text(), '確定')] | "
                    "//input[contains(@value, '關閉') or contains(@value, '確定')]"
                )
                for c_btn in close_btns:
                    if c_btn.is_displayed() and c_btn.is_enabled():
                        try:
                            c_btn.click()
                        except:
                            self.driver.execute_script("arguments[0].click();", c_btn)
                        log_func("  ✓ 已關閉彈窗。")
                        break
            time.sleep(1)
        else:
            log_func("  ℹ️ 未在目前畫面找到『附件歸檔』按鈕，繼續進行分類填寫與存檔。")

        # 5. 【使用者精確需求】關閉附件歸檔分頁後，選擇分類檔號 (q_fsKindno) 與案次號 (q_caseno)
        log_func(f"📌 開始在下拉選單中選擇分類檔號與案次號 ({category_code})...")
        self.driver.switch_to.default_content()
        iframes = self.driver.find_elements(By.TAG_NAME, "iframe")
        for iframe in iframes:
            try:
                src = iframe.get_attribute("src") or ""
                iframe_id = iframe.get_attribute("id") or ""
                if "archive" in src.lower() or "category" in src.lower() or "dialog" in src.lower() or "dtreecontent" in iframe_id.lower():
                    self.driver.switch_to.frame(iframe)
            except:
                pass

        fs_selected = False
        try:
            # 尋找 q_fsKindno 下拉選單
            fs_elements = self.driver.find_elements(By.NAME, "q_fsKindno")
            if not fs_elements:
                fs_elements = self.driver.find_elements(By.ID, "q_fsKindno")
            if not fs_elements:
                fs_elements = self.driver.find_elements(By.XPATH, "//select[contains(@id, 'fs') or contains(@name, 'fs') or contains(@id, 'cls') or contains(@name, 'cls')]")

            if fs_elements:
                select_fs = Select(fs_elements[0])
                # A: 嘗試使用分類代碼 value 選擇
                try:
                    select_fs.select_by_value(category_code)
                    fs_selected = True
                    log_func(f"  ✓ 已成功下拉選擇分類檔號 (value): {category_code}")
                except:
                    pass

                # B: 嘗試遍歷匹配文字
                if not fs_selected:
                    for opt in select_fs.options:
                        opt_val = opt.get_attribute("value") or ""
                        opt_txt = opt.text or ""
                        if category_code in opt_val or category_code in opt_txt:
                            select_fs.select_by_value(opt_val)
                            fs_selected = True
                            log_func(f"  ✓ 已成功下拉選擇匹配之分類檔號: {opt_txt}")
                            break

                time.sleep(0.8)
        except Exception as e_fs:
            log_func(f"  ℹ️ 下拉選擇分類檔號提示: {e_fs}")

        # 若非 Select 下拉選單，備用 input 輸入
        if not fs_selected:
            try:
                cat_inputs = self.driver.find_elements(
                    By.XPATH, 
                    "//input[contains(@id, 'cls') or contains(@name, 'cls') or contains(@id, 'cate') or contains(@name, 'cate') or contains(@placeholder, '分類代碼') or contains(@placeholder, '分類號')]"
                )
                for cat_in in cat_inputs:
                    if cat_in.is_displayed() and cat_in.is_enabled():
                        cat_in.clear()
                        cat_in.send_keys(category_code)
                        log_func(f"  ✓ 已在輸入框填入分類檔號: {category_code}")
                        time.sleep(0.5)
                        break
            except Exception as e_in:
                log_func(f"  ℹ️ 輸入框備用填寫提示: {e_in}")

        # 選擇案次號 (q_caseno)
        try:
            case_elements = self.driver.find_elements(By.ID, "q_caseno")
            if not case_elements:
                case_elements = self.driver.find_elements(By.NAME, "q_caseno")
            if case_elements:
                select_case = Select(case_elements[0])
                if len(select_case.options) > 1:
                    select_case.select_by_index(1)  # 自動選擇第一個有效案次號
                    log_func("  ✓ 已成功選擇案次號 (q_caseno)！")
                elif len(select_case.options) == 1:
                    select_case.select_by_index(0)
                    log_func("  ✓ 已選擇默認案次號 (q_caseno)。")
        except Exception as e_case:
            log_func(f"  ℹ️ 選擇案次號提示: {e_case}")

        # 7. 【使用者核心要求】點擊「確定存檔」/「確定歸檔」按鈕
        log_func("📌 點擊『確定存檔』按鈕送出歸檔作業...")
        self.driver.switch_to.default_content()
        try:
            WebDriverWait(self.driver, 3).until(
                EC.frame_to_be_available_and_switch_to_it((By.ID, "dTreeContent"))
            )
        except:
            pass

        confirm_save_btns = self.driver.find_elements(
            By.XPATH, 
            "//a[contains(text(), '確定存檔') or contains(text(), '確定歸檔') or contains(text(), '確認')] | "
            "//input[contains(@value, '確定存檔') or contains(@value, '確定歸檔') or contains(@value, '確認')] | "
            "//button[contains(text(), '確定存檔') or contains(text(), '確定歸檔') or contains(text(), '確認')]"
        )
        confirm_clicked = False
        for btn in confirm_save_btns:
            if btn.is_displayed() and btn.is_enabled():
                try:
                    btn.click()
                except:
                    self.driver.execute_script("arguments[0].click();", btn)
                confirm_clicked = True
                break

        if not confirm_clicked:
            log_func("  ℹ️ 嘗試直接呼叫系統存檔確認 function...")
            try:
                self.driver.execute_script("if(typeof doConfirmSave !== 'undefined'){ doConfirmSave(); }")
                confirm_clicked = True
            except:
                pass

        log_func("  ✓ 已點擊『確定存檔』！進入憑證與歸檔完成監測階段...")

        # 8. 輪詢與等待歸檔完成（監測憑證簽章或頁面更新）
        log_func("⏳ 正在監測歸檔狀態（若跳出自然人/組織憑證簽章，請在瀏覽器輸入密碼完成簽章）...")
        monitor_start = time.time()
        timeout_seconds = 60

        while time.time() - monitor_start < timeout_seconds:
            time.sleep(3)
            try:
                alert = self.driver.switch_to.alert
                log_func(f"  🔔 系統訊息: {alert.text}")
                alert.accept()
                time.sleep(1)
            except:
                pass

            self.driver.switch_to.default_content()
            try:
                WebDriverWait(self.driver, 2).until(
                    EC.frame_to_be_available_and_switch_to_it((By.ID, "dTreeContent"))
                )
                remaining_cbs = self.driver.find_elements(By.NAME, "ids")
                found_id = any(cb.get_attribute("value") == doc_id for cb in remaining_cbs)
                if not found_id:
                    log_func(f"🎉 [歸檔成功] 公文 ID {doc_id} 已從待結案列表中移除！")
                    return True
            except:
                pass

        log_func(f"⚠️ 公文 ID {doc_id} 已達到單筆監測時間上限 (60秒)，繼續處理下一筆。")
        return True

# ==========================================
# 階段三：Tkinter 主應用程式 GUI (AppGUI)
# ==========================================
class DocArchiverApp:
    def __init__(self, root):
        self.root = root
        self.root.title("📄 電子公文歸檔自動化系統 v1.3")
        self.root.geometry("880x720")
        self.root.minsize(800, 600)

        self.category_mgr = CategoryManager()
        self.engine = SeleniumEngine()

        self.pending_docs = []
        self.combo_vars = {}
        self.combo_widgets = []
        self.is_running = False

        self.setup_ui()
        self.check_version_async()

    def setup_ui(self):
        """設計現代美觀的 GUI 佈局"""
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('TFrame', background='#F4F6F9')
        style.configure('Header.TFrame', background='#1E293B')
        style.configure('Header.TLabel', background='#1E293B', foreground='#FFFFFF', font=('Microsoft JhengHei', 14, 'bold'))
        style.configure('SubHeader.TLabel', background='#1E293B', foreground='#94A3B8', font=('Microsoft JhengHei', 9))
        style.configure('Accent.TButton', font=('Microsoft JhengHei', 10, 'bold'), background='#2563EB', foreground='#FFFFFF')
        style.map('Accent.TButton', background=[('active', '#1D4ED8'), ('disabled', '#CBD5E1')])

        main_container = ttk.Frame(self.root, style='TFrame')
        main_container.pack(fill=tk.BOTH, expand=True)

        header_frame = ttk.Frame(main_container, style='Header.TFrame', padding=(15, 12))
        header_frame.pack(fill=tk.X)

        title_label = ttk.Label(header_frame, text="📄 電子公文歸檔自動化系統", style='Header.TLabel')
        title_label.pack(side=tk.LEFT, anchor=tk.W)

        ver_label = ttk.Label(header_frame, text=f"{CURRENT_VERSION} | 臺北市政府電子公文專用", style='SubHeader.TLabel')
        ver_label.pack(side=tk.LEFT, padx=15, anchor=tk.S)

        self.topmost_var = tk.BooleanVar(value=True)
        self.root.wm_attributes('-topmost', True)
        topmost_chk = tk.Checkbutton(
            header_frame, text="📌 視窗常駐置頂", variable=self.topmost_var,
            command=self.toggle_topmost, bg='#1E293B', fg='#FFFFFF',
            selectcolor='#334155', activebackground='#1E293B', activeforeground='#FFFFFF',
            font=('Microsoft JhengHei', 9)
        )
        topmost_chk.pack(side=tk.RIGHT)

        step1_frame = ttk.LabelFrame(main_container, text=" 第一步：開啟專用 Chrome 與抓取待結案公文 ", padding=10)
        step1_frame.pack(fill=tk.X, padx=15, pady=10)

        step1_desc = ttk.Label(
            step1_frame, 
            text="點擊『啟動專用 Chrome』登入電子公文系統並進入『待結案』頁面後，再點擊『抓取待辦清單』。",
            font=('Microsoft JhengHei', 9), foreground='#475569'
        )
        step1_desc.pack(anchor=tk.W, pady=(0, 8))

        btn_box1 = ttk.Frame(step1_frame)
        btn_box1.pack(fill=tk.X)

        self.btn_launch_chrome = ttk.Button(
            btn_box1, text="🌐 1. 啟動專用 Chrome 瀏覽器", command=self.action_launch_chrome
        )
        self.btn_launch_chrome.pack(side=tk.LEFT, padx=(0, 10))

        self.btn_fetch_docs = ttk.Button(
            btn_box1, text="📥 2. 我已準備好，開始抓取待辦清單！", style='Accent.TButton', command=self.action_fetch_docs
        )
        self.btn_fetch_docs.pack(side=tk.LEFT)

        self.lbl_doc_count = ttk.Label(btn_box1, text="目前尚未抓取資料", font=('Microsoft JhengHei', 9, 'bold'), foreground='#0F172A')
        self.lbl_doc_count.pack(side=tk.RIGHT, padx=5)

        step2_frame = ttk.LabelFrame(main_container, text=" 第二步：設定各筆公文歸檔分類與進行自動歸檔 ", padding=10)
        step2_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=(0, 10))

        batch_bar = ttk.Frame(step2_frame)
        batch_bar.pack(fill=tk.X, pady=(0, 8))

        batch_lbl = ttk.Label(batch_bar, text="⚡ 快速統一設定：", font=('Microsoft JhengHei', 9, 'bold'))
        batch_lbl.pack(side=tk.LEFT)

        self.batch_combo_var = tk.StringVar()
        category_options = list(self.category_mgr.category_map.keys())
        self.batch_combo = ttk.Combobox(
            batch_bar, textvariable=self.batch_combo_var, values=category_options,
            state="readonly", width=38, font=('Microsoft JhengHei', 9)
        )
        if category_options:
            self.batch_combo.current(0)
        self.batch_combo.pack(side=tk.LEFT, padx=5)

        btn_apply_all = ttk.Button(batch_bar, text="套用至下方全部公文", command=self.action_apply_batch_category)
        btn_apply_all.pack(side=tk.LEFT, padx=5)

        list_container = ttk.Frame(step2_frame)
        list_container.pack(fill=tk.BOTH, expand=True)

        self.canvas = tk.Canvas(list_container, bg='#FFFFFF', highlightthickness=1, highlightbackground='#CBD5E1')
        scrollbar = ttk.Scrollbar(list_container, orient="vertical", command=self.canvas.yview)
        
        self.scrollable_frame = ttk.Frame(self.canvas)
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=scrollbar.set)

        self.canvas.bind_all("<MouseWheel>", lambda e: self.canvas.yview_scroll(int(-1 * (e.delta / 120)), "units"))

        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        bottom_frame = ttk.Frame(main_container, padding=(15, 0, 15, 10))
        bottom_frame.pack(fill=tk.X)

        self.btn_start_archive = ttk.Button(
            bottom_frame, text="⚡ 確認分類，開始批次自動歸檔！", style='Accent.TButton', command=self.action_start_archive
        )
        self.btn_start_archive.pack(fill=tk.X, ipady=5, pady=(0, 8))

        log_lbl = ttk.Label(bottom_frame, text="📋 即時執行日誌 (Live Execution Log):", font=('Microsoft JhengHei', 9, 'bold'))
        log_lbl.pack(anchor=tk.W, pady=(0, 2))

        self.log_area = scrolledtext.ScrolledText(
            bottom_frame, height=7, font=('Consolas', 9), bg='#0F172A', fg='#E2E8F0', insertbackground='white'
        )
        self.log_area.pack(fill=tk.X)
        self.log("💡 系統初始化完成。請遵循步驟 1 啟動 Chrome 並抓取清單。")

    def toggle_topmost(self):
        self.root.wm_attributes('-topmost', self.topmost_var.get())

    def log(self, message):
        timestamp = time.strftime("[%H:%M:%S] ")
        full_msg = timestamp + message + "\n"
        
        def append_log():
            self.log_area.insert(tk.END, full_msg)
            self.log_area.see(tk.END)
            
        if threading.current_thread() is threading.main_thread():
            append_log()
        else:
            self.root.after(0, append_log)

    def check_version_async(self):
        def work():
            latest, url = check_for_updates()
            if latest:
                self.log(f"🔔 發現 GitHub 最新版本 {latest}！建議前往更新: {url}")
        threading.Thread(target=work, daemon=True).start()

    def action_launch_chrome(self):
        ok, msg = self.engine.launch_chrome_process()
        if ok:
            messagebox.showinfo("成功", f"{msg}\n\n請在開啟的 Chrome 視窗中：\n1. 登入電子公文系統\n2. 切換至『待結案』畫面\n3. 完成後回到此視窗點擊『抓取待辦清單』。")
            self.log("🌐 已送出專用 Chrome 啟動指令。")
        else:
            messagebox.showerror("錯誤", msg)
            self.log(f"❌ {msg}")

    def action_fetch_docs(self):
        if self.is_running:
            return

        self.btn_fetch_docs.config(state="disabled")
        self.log("🔄 開始讀取 Chrome 待辦清單，請稍候...")

        def thread_target():
            try:
                docs = self.engine.fetch_pending_docs(log_func=self.log)
                self.root.after(0, lambda: self.render_pending_docs(docs))
            except Exception as e:
                err_msg = str(e)
                self.log(f"❌ 讀取清單發生錯誤: {err_msg}")
                self.root.after(0, lambda: messagebox.showerror("錯誤", f"無法抓取公文清單:\n{err_msg}"))
            finally:
                self.root.after(0, lambda: self.btn_fetch_docs.config(state="normal"))

        threading.Thread(target=thread_target, daemon=True).start()

    def render_pending_docs(self, docs):
        self.pending_docs = docs
        self.combo_vars.clear()
        self.combo_widgets.clear()

        for child in self.scrollable_frame.winfo_children():
            child.destroy()

        if not docs:
            self.lbl_doc_count.config(text="未找到待辦公文", foreground='#DC2626')
            messagebox.showinfo("提示", "目前頁面上未找到任何待辦公文。\n請確認 Chrome 是否已登入並切換至『待結案』頁面。")
            self.log("ℹ️ 未抓取到待辦公文。")
            return

        self.lbl_doc_count.config(text=f"已抓取到 {len(docs)} 筆公文", foreground='#16A34A')
        self.log(f"🎉 成功抓取 {len(docs)} 筆待辦公文！請為各筆設定歸檔分類。")

        category_options = list(self.category_mgr.category_map.keys())

        header_row = ttk.Frame(self.scrollable_frame, padding=(5, 5))
        header_row.pack(fill=tk.X, expand=True)
        ttk.Label(header_row, text="序號 / 公文 ID", font=('Microsoft JhengHei', 9, 'bold'), width=22).pack(side=tk.LEFT)
        ttk.Label(header_row, text="公文主旨摘要", font=('Microsoft JhengHei', 9, 'bold'), width=45).pack(side=tk.LEFT, padx=5)
        ttk.Label(header_row, text="選擇歸檔分類與案號", font=('Microsoft JhengHei', 9, 'bold')).pack(side=tk.LEFT)

        ttk.Separator(self.scrollable_frame, orient='horizontal').pack(fill=tk.X, pady=2)

        for idx, doc in enumerate(docs, 1):
            row_bg = '#F8FAFC' if idx % 2 == 0 else '#FFFFFF'
            row_frame = tk.Frame(self.scrollable_frame, bg=row_bg, padx=5, pady=4)
            row_frame.pack(fill=tk.X, expand=True)

            doc_id = doc["id"]
            title = doc["title"]

            lbl_id = tk.Label(
                row_frame, text=f"{idx}. {doc_id}", font=('Consolas', 9),
                bg=row_bg, anchor="w", width=22
            )
            lbl_id.pack(side=tk.LEFT)

            lbl_title = tk.Label(
                row_frame, text=title, font=('Microsoft JhengHei', 9),
                bg=row_bg, anchor="w", width=45, wraplength=340, justify="left"
            )
            lbl_title.pack(side=tk.LEFT, padx=5)

            var = tk.StringVar()
            combo = ttk.Combobox(
                row_frame, textvariable=var, values=category_options,
                state="readonly", width=36, font=('Microsoft JhengHei', 9)
            )
            if category_options:
                combo.current(0)
            combo.pack(side=tk.LEFT, padx=5)

            self.combo_vars[doc_id] = var
            self.combo_widgets.append(combo)

        self.scrollable_frame.update_idletasks()
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def action_apply_batch_category(self):
        target_val = self.batch_combo_var.get()
        if not target_val:
            return
        
        count = 0
        for var in self.combo_vars.values():
            var.set(target_val)
            count += 1
            
        self.log(f"⚡ 已將分類『{target_val}』快速套用至全部 {count} 筆公文。")

    def action_start_archive(self):
        if self.is_running:
            return

        if not self.pending_docs:
            messagebox.showwarning("警告", "請先完成步驟一『抓取待辦清單』！")
            return

        tasks = []
        for doc in self.pending_docs:
            doc_id = doc["id"]
            title = doc["title"]
            selected_disp = self.combo_vars[doc_id].get()
            cat_code = self.category_mgr.category_map.get(selected_disp)

            if cat_code is not None:
                tasks.append({
                    "id": doc_id,
                    "title": title,
                    "code": cat_code,
                    "disp": selected_disp
                })

        if not tasks:
            messagebox.showinfo("提示", "沒有需要自動歸檔的項目 (所有項目均設定為手動處理)。")
            return

        confirm = messagebox.askyesno(
            "確認開始歸檔", 
            f"即將開始為 {len(tasks)} 筆公文執行全自動歸檔！\n\n"
            "歸檔過程中程式將會自動執行：\n"
            " 1. 勾選公文並點擊存查\n"
            " 2. 輸入分類號與案號\n"
            " 3. 點擊『附件歸檔』按鈕\n"
            " 4. 自動關閉『附件歸檔』彈窗\n"
            " 5. 點擊『確定存檔』送出\n\n"
            "請確保專用 Chrome 瀏覽器保持在最前或可存取狀態。\n確定要開始嗎？"
        )
        if not confirm:
            return

        self.is_running = True
        self.btn_start_archive.config(state="disabled", text="⏳ 自動歸檔進行中...")
        self.log("🚀 ================= 開始批次自動歸檔任務 =================")

        def thread_worker():
            success_count = 0
            fail_count = 0

            for idx, task in enumerate(tasks, 1):
                self.log(f"\n📌 進度: ({idx}/{len(tasks)})")
                try:
                    res = self.engine.process_single_doc_archive(
                        doc_id=task["id"],
                        doc_title=task["title"],
                        category_code=task["code"],
                        log_func=self.log
                    )
                    if res:
                        success_count += 1
                    else:
                        fail_count += 1
                except Exception as e:
                    fail_count += 1
                    self.log(f"❌ 處理公文 {task['id']} 時發生異常: {e}")
                    self.log(traceback.format_exc())

                time.sleep(2)

            self.log("\n==================================================")
            self.log(f"🏁 批次歸檔任務結束！成功: {success_count} 筆，失敗/跳過: {fail_count} 筆。")
            self.log("==================================================")

            def on_finish():
                self.is_running = False
                self.btn_start_archive.config(state="normal", text="⚡ 確認分類，開始批次自動歸檔！")
                messagebox.showinfo(
                    "完成", 
                    f"批次自動歸檔任務完成！\n\n成功處理: {success_count} 筆\n失敗或跳過: {fail_count} 筆\n\n詳情請查看下方日誌視窗。"
                )

            self.root.after(0, on_finish)

        threading.Thread(target=thread_worker, daemon=True).start()

# ==========================================
# 主程式進入點
# ==========================================
def main():
    root = tk.Tk()
    app = DocArchiverApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
