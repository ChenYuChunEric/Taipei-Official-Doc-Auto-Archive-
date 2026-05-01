import time
import tkinter as tk
from tkinter import ttk, messagebox
import traceback
import subprocess
import os
import sys
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select

# ==========================================
# 階段零：環境設定與「極簡版」外部設定檔讀取
# ==========================================

# 這是內建的「大腦辭典」，包含了系統內所有的分類代碼與完整名稱
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

def load_categories():
    """讀取 categories.txt 並自動比對大腦辭典補全名稱"""
    if getattr(sys, 'frozen', False):
        base_path = os.path.dirname(sys.executable)
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
        
    config_file = os.path.join(base_path, 'categories.txt')
    category_map = {}
    
    # 給同事的極簡版預設代碼範例
    default_codes = [
        "03750401", "03750402", "03750403", 
        "03750404", "03750501", "03750502"
    ]

    # 如果沒有檔案，幫他產生一個只有數字的記事本
    if not os.path.exists(config_file):
        try:
            with open(config_file, 'w', encoding='utf-8') as f:
                f.write("# ==========================================\n")
                f.write("# 電子公文歸檔小幫手 - 專屬分類設定檔\n")
                f.write("# 請在下方輸入您常用的「分類代碼數字」即可 (一行一個)\n")
                f.write("# 程式啟動時會自動幫您轉換成完整的分類名稱！\n")
                f.write("# ==========================================\n\n")
                for code in default_codes:
                    f.write(f"{code}\n")
            
            # 使用預設代碼建構 UI 顯示表
            for code in default_codes:
                display_name = MASTER_CATEGORIES.get(code, f"{code} (未知分類代碼)")
                category_map[display_name] = code
            print(f"📁 已自動產生預設分類檔：categories.txt")
        except Exception as e:
            pass

    else:
        # 如果檔案存在，讀取裡面的數字
        try:
            with open(config_file, 'r', encoding='utf-8') as f:
                for line in f:
                    code = line.strip()
                    # 略過空白或註解
                    if not code or code.startswith("#"):
                        continue
                    
                    # ✨ 核心魔法：拿數字去查大腦辭典
                    display_name = MASTER_CATEGORIES.get(code, f"{code} (未知分類代碼)")
                    category_map[display_name] = code
                    
            print(f"📁 成功讀取 categories.txt 並完成名稱對應！")
        except Exception as e:
            print(f"讀取設定檔發生錯誤：{e}")

    # 永遠把「跳過」選項加在最下面
    category_map["暫不處理 (跳過此份公文)"] = None
    return category_map

# 執行讀取並建立 UI 下拉選單要用的字典
CATEGORY_MAP = load_categories()

# 全域變數準備
pending_docs = []
user_selections = {}
combo_vars = {}
driver = None

# ==========================================
# 階段一：自動開啟 Chrome 專屬分身 (同事共享升級版)
# ==========================================
def launch_chrome():
    # 自動尋找 Chrome 執行檔的可能路徑 (兼顧 64 位元與 32 位元系統)
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
        messagebox.showerror("錯誤", "找不到 Chrome 瀏覽器！請確認同事電腦是否有安裝 Chrome。")
        return False

    # 💡 動態抓取當前電腦使用者的「家目錄」 (例如 C:\Users\王小明)
    user_home = os.path.expanduser("~")
    # 為該同事建立專屬的自動化資料夾
    user_data_dir = os.path.join(user_home, "SeleniumChromeProfile")

    print(f"🚀 正在啟動專屬辦公 Chrome...")
    print(f"📁 專屬暫存資料夾建立於：{user_data_dir}")
    
    try:
        # 使用 subprocess 自動在背景開啟 Chrome
        subprocess.Popen([
            chrome_path, 
            "--remote-debugging-port=9222", 
            f"--user-data-dir={user_data_dir}"
        ])
        return True
    except Exception as e:
        messagebox.showerror("啟動失敗", f"無法啟動 Chrome: {e}")
        return False

# ==========================================
# UI 邏輯控制中心
# ==========================================
def step1_fetch_docs():
    """當使用者按下第一階段的「準備好」按鈕時執行"""
    global driver, pending_docs
    
    btn_start_fetch.config(text="連線中，請稍候...", state="disabled")
    root.update()

    try:
        # 1. 接管已經開啟的 Chrome
        chrome_options = Options()
        chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
        driver = webdriver.Chrome(options=chrome_options)
        
        # 2. 抓取待結案清單
        driver.switch_to.default_content()
        WebDriverWait(driver, 5).until(
            EC.frame_to_be_available_and_switch_to_it((By.ID, "dTreeContent"))
        )

        rows = driver.find_elements(By.XPATH, "//tbody[@id='listTBODY']/tr")
        pending_docs.clear()
        
        for row in rows:
            checkbox = row.find_element(By.NAME, "ids")
            doc_id = checkbox.get_attribute("value")
            title_element = row.find_element(By.XPATH, ".//td[@data-th='主旨摘要']//span[@id='mainSpan']")
            pending_docs.append({"id": doc_id, "title": title_element.text})

        if not pending_docs:
            messagebox.showinfo("提示", "目前畫面上沒有待辦公文！請確認是否停留在「待結案」清單。")
            btn_start_fetch.config(text="我已準備好，開始抓取清單！", state="normal")
            return

        # 如果成功抓到，切換到第二階段 UI
        show_step2_ui()

    except Exception as e:
        messagebox.showerror("連線失敗", f"無法抓取公文，請確認是否已進入「待結案」畫面。\n錯誤訊息：{type(e).__name__}")
        btn_start_fetch.config(text="我已準備好，開始抓取清單！", state="normal")

def show_step2_ui():
    """顯示第二階段：分類下拉選單"""
    # 清空第一階段的畫面元件
    for widget in frame_step1.winfo_children():
        widget.destroy()
    frame_step1.pack_forget()

    # 顯示第二階段的畫面元件
    frame_step2.pack(fill="both", expand=True)

    # 建立滾動視窗
    canvas = tk.Canvas(frame_step2)
    scrollbar = ttk.Scrollbar(frame_step2, orient="vertical", command=canvas.yview)
    scrollable_frame = ttk.Frame(canvas)

    scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    canvas.pack(side="top", fill="both", expand=True, padx=10, pady=10)
    scrollbar.pack(side="right", fill="y")

    lbl_info = ttk.Label(scrollable_frame, text=f"✅ 成功抓取 {len(pending_docs)} 筆公文，請選擇歸檔分類：", font=("微軟正黑體", 12, "bold"))
    lbl_info.grid(row=0, column=0, columnspan=2, pady=10, sticky="w")

    # 動態產生選單
    global combo_vars
    for idx, doc in enumerate(pending_docs):
        ttk.Label(scrollable_frame, text=f"{idx+1}. {doc['title']}", wraplength=450).grid(row=idx+1, column=0, padx=5, pady=10, sticky="w")
        var = tk.StringVar()
        combo = ttk.Combobox(scrollable_frame, textvariable=var, values=list(CATEGORY_MAP.keys()), width=35, state="readonly")
        combo.current(0)
        combo.grid(row=idx+1, column=1, padx=5, pady=10)
        combo_vars[doc["id"]] = var

    btn_execute = ttk.Button(frame_step2, text="確認分類，開始批次自動歸檔！", command=start_batch_process)
    btn_execute.pack(pady=15)

def start_batch_process():
    """收集分類結果並關閉 UI，進入自動化流程"""
    for doc in pending_docs:
        selected_name = combo_vars[doc["id"]].get()
        user_selections[doc["id"]] = CATEGORY_MAP[selected_name]
    root.destroy() # 關閉 UI 視窗，讓後面的 Python 程式碼繼續執行

# ==========================================
# 啟動主程式與 UI 介面
# ==========================================
try:
    # 1. 先啟動 Chrome
    if not launch_chrome():
        exit()

    # 2. 建立 Tkinter 視窗
    root = tk.Tk()
    root.title("電子公文歸檔自動化系統")
    root.geometry("850x450")
    root.attributes('-topmost', True) # 視窗置頂，方便操作

    # 第一階段 Frame (等待準備)
    frame_step1 = ttk.Frame(root)
    frame_step1.pack(fill="both", expand=True)

    lbl_title = ttk.Label(frame_step1, text="🌐 專屬瀏覽器已開啟！", font=("微軟正黑體", 16, "bold"))
    lbl_title.pack(pady=30)
    
    instructions = (
        "請按照以下步驟操作：\n\n"
        "1. 在剛開啟的 Chrome 中，登入您的電子公文系統。\n"
        "2. 插入憑證完成驗證，並正常批閱公文。\n"
        "3. 當準備要歸檔時，請點擊左側選單進入「待結案」畫面。\n"
        "4. 確認畫面上顯示待辦公文後，點擊下方按鈕。"
    )
    lbl_desc = ttk.Label(frame_step1, text=instructions, font=("微軟正黑體", 12), justify="left")
    lbl_desc.pack(pady=10)

    btn_start_fetch = ttk.Button(frame_step1, text="我已準備好，開始抓取清單！", command=step1_fetch_docs)
    btn_start_fetch.pack(pady=30)

    # 第二階段 Frame (分類選單，先隱藏)
    frame_step2 = ttk.Frame(root)

    # 啟動 UI 監聽
    root.mainloop()

    # ==========================================
    # 階段三：開始批次自動化歸檔執行
    # (只有在 UI 被正常關閉後才會走到這裡)
    # ==========================================
    if not user_selections:
        print("未執行批次歸檔，程式結束。")
        exit()

    print("\n🚀 開始批次歸檔作業...")

    for doc in pending_docs:
        cat_val = user_selections.get(doc["id"])
        if not cat_val:
            print(f"⏩ 略過公文：{doc['title'][:15]}...")
            continue

        print(f"⏳ 正在處理：{doc['title'][:15]}...")

        try:
            driver.switch_to.default_content()
            WebDriverWait(driver, 15).until(
                EC.frame_to_be_available_and_switch_to_it((By.ID, "dTreeContent"))
            )
            WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.ID, "listTBODY"))
            )
            time.sleep(1)

            print("  -> 正在智能判斷並勾選目標公文...")
            smart_check_script = f"""
                var targetId = '{doc['id']}';
                var checkboxes = document.getElementsByName('ids');
                for(var i=0; i<checkboxes.length; i++) {{
                    var cb = checkboxes[i];
                    if(cb.value === targetId) {{
                        if(!cb.checked) {{ $(cb).iCheck('check'); cb.checked = true; }}
                    }} else {{
                        if(cb.checked) {{ $(cb).iCheck('uncheck'); cb.checked = false; }}
                    }}
                }}
            """
            driver.execute_script(smart_check_script)
            time.sleep(2)

            print("  -> 正在尋找並點擊「存查」按鈕...")
            success_transition = False
            
            for attempt in range(3):
                archive_btns = driver.find_elements(By.XPATH, "//input[@value='存查']")
                for btn in archive_btns:
                    if btn.is_displayed():
                        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn)
                        time.sleep(0.5)
                        try:
                            btn.click() 
                        except:
                            driver.execute_script("arguments[0].click();", btn) 
                        break 
                
                try:
                    alert = WebDriverWait(driver, 2).until(EC.alert_is_present())
                    error_msg = alert.text
                    alert.accept()
                    print(f"❌ 系統警告：{error_msg}")
                    break 
                except:
                    pass
                
                try:
                    WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.NAME, "q_fsKindno")))
                    success_transition = True
                    break 
                except:
                    continue
            
            if not success_transition:
                print("❌ 無法順利進入存查頁面，將略過此筆公文。")
                continue
            
            print("  -> 正在填寫歸檔分類...")
            fs_select_element = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.NAME, "q_fsKindno"))
            )
            Select(fs_select_element).select_by_value(cat_val)

            WebDriverWait(driver, 10).until(
                lambda d: len(Select(d.find_element(By.ID, "q_caseno")).options) > 1
            )
            Select(driver.find_element(By.ID, "q_caseno")).select_by_index(1)

            print("  -> 正在送出歸檔...")
            submit_btn = driver.find_element(By.ID, "updateSubmit")
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", submit_btn)
            time.sleep(0.5)
            try:
                submit_btn.click()
            except:
                driver.execute_script("arguments[0].click();", submit_btn)

            try:
                alert = WebDriverWait(driver, 3).until(EC.alert_is_present())
                alert.accept() 
            except:
                pass 

            print("  -> 等待硬體驗證與系統跳轉 (時間不固定，請稍候)...")
            
            wait_time = 0
            success_archived = False
            
            while wait_time < 90:
                time.sleep(5)
                wait_time += 5
                
                try:
                    driver.switch_to.default_content()
                    driver.switch_to.frame("dTreeContent")
                    
                    # 確保已回到含有待結案清單的畫面
                    driver.find_element(By.ID, "listTBODY")
                    
                    # 抓取畫面上剩餘的公文 ID
                    new_checkboxes = driver.find_elements(By.NAME, "ids")
                    new_ids = [cb.get_attribute("value") for cb in new_checkboxes]
                    
                    # 若當前處理的公文ID已不在清單中，視為歸檔完成並跳轉回清單
                    if doc['id'] not in new_ids:
                        success_archived = True
                        break
                except Exception:
                    # 發生錯誤代表頁面可能正在跳轉或載入中，忽略並繼續等待
                    pass
            
            if success_archived:
                print(f"✅ 完成歸檔：{doc['title'][:15]}...")
            else:
                print(f"⚠️ 等待超時或未偵測到變化，強制接續下一筆...")
                
            time.sleep(1.5)

        except Exception as inner_e:
            print(f"❌ 處理此公文時發生錯誤：{type(inner_e).__name__}")
            print("將嘗試重置畫面以處理下一筆...")
            try:
                driver.switch_to.default_content()
                driver.find_element(By.ID, "gotoFlow").click()
                time.sleep(2)
            except:
                pass

    print("\n🎉 所有批次歸檔作業已經順利完成！")

except Exception as main_e:
    print("\n" + "="*40)
    print("🚨 程式發生致命錯誤：")
    print("="*40)
    traceback.print_exc()
    print("="*40)

finally:
    if driver:
        # 結束時詢問是否要關閉 Chrome，如果不要關閉可以把這行註解掉
        pass
    input("\n執行結束，請按下 Enter 鍵關閉這個視窗...")