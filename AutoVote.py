import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
import threading
import sys
import os
import time
import io
import re
import datetime
import shutil
import hashlib
import base64
import subprocess
import concurrent.futures
import threading
import uuid
import json
import requests
import subprocess
import urllib.parse
from __version__ import CURRENT_VERSION


VERSION_INFO_URL = "https://raw.githubusercontent.com/marslin1027-sudo/AutoVoteApp/refs/heads/main/version.json"
from PIL import Image

# 強化加密所需模組
from cryptography.fernet import Fernet
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
from cryptography.hazmat.backends import default_backend

# Selenium Imports
from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.common.exceptions import StaleElementReferenceException

# --- 自定義異常 ---
class SystemMaintenanceError(Exception):
    pass

class LoginTimeoutError(Exception): 
    pass

# --- 全域變數初始化 ---
debug = 0
driver = None
voteinfolist = {}
base_path = "./screenshots/"
shareholderIDs = []
saved_sites = {} # <--- 新增：用來儲存自訂網址的字典
browser_choice = "Edge" # <--- 新增：預設瀏覽器

# 新增：速率分離與視窗記憶
vote_speed = 2.0  
shot_speed = 0.5
main_window_geom = ""
disc_window_geom = ""

screenshot_mode = 1 
manual_vote = False
default_vote = "abstain"
login_type = "券商網路下單憑證" 
accept_list = []
opposite_list = []
abstain_list = []
disclaimer_agreed = False
join_draw = False # <--- 新增：是否參加抽獎的預設值

session_results = {}
user_name_map = {}
execution_logs = []
ignore_update_until = 0.0 # <--- 新增：紀錄暫停自動檢查更新的時間戳記
last_selected_site = "【不開啟任何網頁】" # <--- 新增：記憶上次選擇的網頁

os.environ["QTWEBENGINE_CHROMIUM_FLAGS"] = "--log-level=3"

# --- 更嚴謹的硬體金鑰衍生演算法 (PBKDF2HMAC) ---
def get_hw_key():
    try:
        cmd = 'wmic baseboard get serialnumber'
        output = subprocess.check_output(cmd, shell=True).decode(errors='ignore')
        hw_id = output.split('\n')[1].strip()
        
        if not hw_id or "To be filled" in hw_id or "None" in hw_id:
            cmd = 'wmic cpu get processorid'
            output = subprocess.check_output(cmd, shell=True).decode(errors='ignore')
            hw_id = output.split('\n')[1].strip()
            
        if not hw_id or "To be filled" in hw_id or "None" in hw_id:
             hw_id = str(uuid.getnode())
             
    except:
        try: hw_id = str(uuid.getnode())
        except: hw_id = "Default_Fallback_Seed_12345"
        
    # 加入 Salt 與 10萬次迭代增加破解難度
    salt = b'TDCC_AutoVote_Secure_Salt_V4'
    kdf = PBKDF2HMAC(
        algorithm=hashes.SHA256(),
        length=32,
        salt=salt,
        iterations=100000,
        backend=default_backend()
    )
    return base64.urlsafe_b64encode(kdf.derive(hw_id.encode()))

cipher = Fernet(get_hw_key())

def encrypt_data(data_str):
    if not data_str: return ""
    return cipher.encrypt(data_str.encode()).decode()

def decrypt_data(encrypted_str):
    if not encrypted_str: return ""
    try:
        return cipher.decrypt(encrypted_str.encode()).decode()
    except:
        return "" 

def get_anonymous_dirname(user_id):
    if not user_id: return "unknown"
    return hashlib.sha256(user_id.encode()).hexdigest()

# --- 路徑處理函數 ---
def get_app_data_path():
    app_data = os.getenv('APPDATA')
    if not app_data:
        app_data = os.path.expanduser("~")
    config_dir = os.path.join(app_data, "TDCC_AutoVote_Configs")
    if not os.path.exists(config_dir):
        os.makedirs(config_dir)
    return config_dir

CONFIG_DIR = get_app_data_path()

import urllib.parse # 確保最上方有這個模組，或是直接讓它在函式裡被呼叫也可以

def check_for_updates(auto=False):
    # 啟動背景執行緒來檢查網路，避免卡死主視窗 UI
    threading.Thread(target=_check_for_updates_thread, args=(auto,), daemon=True).start()

def _check_for_updates_thread(auto):
    global ignore_update_until
    
    # 如果是系統自動檢查，且目前時間還在「7天免打擾」期間內，則直接略過
    if auto and time.time() < ignore_update_until:
        return

    try:
        timestamp = int(time.time())
        url_with_timestamp = f"{VERSION_INFO_URL}?t={timestamp}"
        
        response = requests.get(url_with_timestamp, timeout=5)
        data = response.json()
        remote_version = data.get("version")
        download_url = data.get("url")
        update_notes = data.get("notes", "無詳細更新說明")

        def parse_version(v):
            return [int(x) for x in v.split('.')]

        if parse_version(remote_version) > parse_version(CURRENT_VERSION):
            # 必須回到主執行緒 (Main Thread) 才能建立 Tkinter 視窗！
            app.after(0, lambda: _show_update_window(remote_version, download_url, update_notes, auto))
        else:
            # 只有手動點擊檢查時，才跳出「已是最新版本」的提示
            if not auto:
                app.after(0, lambda: messagebox.showinfo("檢查更新", f"目前已是最新版本 (v{CURRENT_VERSION})"))
    except Exception as e:
        if not auto:
            app.after(0, lambda: messagebox.showerror("更新失敗", f"無法檢查更新，請確認網路連線。\n({e})"))

def _show_update_window(remote_version, download_url, update_notes, auto):
    # --- 建立自定義的更新通知視窗 ---
    update_win = tk.Toplevel()
    update_win.title("軟體更新通知")
    update_win.attributes('-topmost', True) # 視窗置頂
    
    # 設定一個最小寬高，避免內容太少時視窗縮得太小
    update_win.minsize(400, 200)

    msg = (
        f"🎉 發現新版本！\n\n"
        f"目前版本: v{CURRENT_VERSION}\n"
        f"最新版本: v{remote_version}\n\n"
        f"【本次更新內容】\n"
        f"{update_notes}\n"
    )
    
    # 用一個 Frame 把所有東西包起來，方便控制邊距
    content_frame = ttk.Frame(update_win, padding=20)
    content_frame.pack(fill="both", expand=True)
    
    # Label 會根據內容的多寡自動把高度撐開，wraplength 控制寬度到多少時自動換行
    msg_label = tk.Label(content_frame, text=msg, font=("Microsoft JhengHei", 10), justify="left", wraplength=360)
    msg_label.pack(fill="both", expand=True)
    
    ignore_var = tk.BooleanVar(value=False)
    
    # 只有自動檢查觸發時，才顯示「7天內不再檢查」的勾選框
    if auto:
        ttk.Checkbutton(content_frame, text="7天內不再自動檢查更新", variable=ignore_var).pack(pady=(10, 0))
    
    btn_frame = ttk.Frame(content_frame)
    btn_frame.pack(pady=(20, 0))
    
    def on_yes():
        update_win.destroy()
        perform_update(download_url)
        
    def on_no():
        if auto and ignore_var.get():
            global ignore_update_until
            # 紀錄 7 天後的時間戳記
            ignore_update_until = time.time() + 7 * 24 * 60 * 60
            try: app.save_config() # 呼叫主程式存檔
            except: pass
        update_win.destroy()
        
    ttk.Button(btn_frame, text="立即下載更新", command=on_yes, style='Action.TButton').pack(side="left", padx=10, ipadx=10)
    ttk.Button(btn_frame, text="稍後再說", command=on_no).pack(side="left", padx=10, ipadx=10)

    # ==========================================
    # 🌟 動態計算視窗大小與置中魔法
    # ==========================================
    # 先讓系統在背景排版，計算出各元件實際需要的長寬
    update_win.update_idletasks()
    
    # 取得自動撐開後的實際寬度與高度
    req_w = update_win.winfo_reqwidth()
    req_h = update_win.winfo_reqheight()
    
    # 計算置中座標
    x = (update_win.winfo_screenwidth() // 2) - (req_w // 2)
    y = (update_win.winfo_screenheight() // 2) - (req_h // 2)
    
    # 套用動態計算出的大小與置中位置
    update_win.geometry(f'{req_w}x{req_h}+{x}+{y}')

            
def perform_update(download_url):
    try:
        current_exe_path = os.path.abspath(sys.executable)
        current_dir = os.path.dirname(current_exe_path)
        current_exe_name = os.path.basename(current_exe_path)
        
        if not getattr(sys, 'frozen', False):
            messagebox.showinfo("提示", "目前在開發環境(py檔)，請打包成 EXE 後再測試自動更新。")
            return

        new_exe_name = urllib.parse.unquote(download_url.split('/')[-1])
        temp_exe_name = "Update_Temp.exe"
        temp_exe_path = os.path.join(current_dir, temp_exe_name)
        bat_path = os.path.join(current_dir, "updater.bat")

        # ==========================================
        # 🎨 建立獨立的下載進度條視窗
        # ==========================================
        progress_win = tk.Toplevel()
        progress_win.title("軟體更新中")
        progress_win.geometry("350x160")
        progress_win.attributes('-topmost', True) # 視窗置頂
        
        # 讓進度視窗置中於螢幕
        progress_win.update_idletasks()
        w = progress_win.winfo_width()
        h = progress_win.winfo_height()
        x = (progress_win.winfo_screenwidth() // 2) - (w // 2)
        y = (progress_win.winfo_screenheight() // 2) - (h // 2)
        progress_win.geometry(f'{w}x{h}+{x}+{y}')

        tk.Label(progress_win, text="🚀 正在為您下載最新版本...", font=("Microsoft JhengHei", 11, "bold")).pack(pady=(15, 5))
        tk.Label(progress_win, text="下載完成前請勿關閉程式", font=("Microsoft JhengHei", 9), fg="#d9534f").pack(pady=(0, 10))
        
        progress_var = tk.DoubleVar()
        progress_bar = ttk.Progressbar(progress_win, variable=progress_var, maximum=100, length=280)
        progress_bar.pack(pady=5)
        
        status_label = tk.Label(progress_win, text="準備連線...", font=("Consolas", 9))
        status_label.pack(pady=5)
        progress_win.update()

        # ==========================================
        # 🚀 極速下載引擎 (多執行緒 + UI 連動)
        # ==========================================
        head_req = requests.head(download_url, allow_redirects=True)
        total_size = int(head_req.headers.get('content-length', 0))
        downloaded_size = 0
        lock = threading.Lock()

        if total_size > 0:
            threads_count = 8
            chunk_size = total_size // threads_count
            
            def download_range(start, end):
                nonlocal downloaded_size
                headers = {'Range': f'bytes={start}-{end}'}
                resp = requests.get(download_url, headers=headers, stream=True)
                data = bytearray()
                for chunk in resp.iter_content(chunk_size=65536):
                    if chunk:
                        data.extend(chunk)
                        # 安全地累加已下載的量
                        with lock:
                            downloaded_size += len(chunk)
                return data
            
            with concurrent.futures.ThreadPoolExecutor(max_workers=threads_count) as executor:
                futures = []
                for i in range(threads_count):
                    start = i * chunk_size
                    end = start + chunk_size - 1 if i < threads_count - 1 else total_size - 1
                    futures.append(executor.submit(download_range, start, end))
                
                # 主執行緒負責不斷刷新 UI，直到下載完畢
                while any(not f.done() for f in futures):
                    with lock:
                        current_dl = downloaded_size
                    percent = (current_dl / total_size) * 100
                    progress_var.set(percent)
                    status_label.config(text=f"已下載: {current_dl/(1024*1024):.1f} MB / {total_size/(1024*1024):.1f} MB ({percent:.1f}%)")
                    progress_win.update()
                    time.sleep(0.05) # 稍微休息避免 UI 卡死

                # 組合碎片存檔
                with open(temp_exe_path, 'wb') as f:
                    for future in futures:
                        f.write(future.result())
        else:
            # 防呆：如果伺服器不給檔案大小，使用單線程 + 跑馬燈模式
            with requests.get(download_url, stream=True) as r:
                r.raise_for_status()
                with open(temp_exe_path, 'wb') as f:
                    for chunk in r.iter_content(chunk_size=1024*1024):
                        f.write(chunk)
                        downloaded_size += len(chunk)
                        progress_bar.config(mode='indeterminate')
                        progress_bar.step(10)
                        status_label.config(text=f"已下載: {downloaded_size/(1024*1024):.1f} MB")
                        progress_win.update()

        # ==========================================
        # 🔄 下載完成，產生 BAT 腳本並重啟
        # ==========================================
        status_label.config(text="✅ 下載完成！正在重新啟動...")
        progress_var.set(100)
        progress_bar.config(mode='determinate')
        progress_win.update()
        time.sleep(1) # 讓使用者看一眼「下載完成」的字眼
        progress_win.destroy()

        bat_content = f"""@echo off
chcp 65001 > nul
setlocal enabledelayedexpansion

cd /d "{current_dir}"
timeout /t 2 /nobreak > nul

set "retry=0"
:delete_old
if exist "{current_exe_name}" (
    del /f /q "{current_exe_name}"
    if exist "{current_exe_name}" (
        set /a retry+=1
        if !retry! lss 10 (
            timeout /t 1 /nobreak > nul
            goto delete_old
        ) else (
            goto end
        )
    )
)

if exist "{new_exe_name}" del /f /q "{new_exe_name}"
ren "{temp_exe_name}" "{new_exe_name}"
start "" "{new_exe_name}"

:end
del "%~f0"
"""
        with open(bat_path, "w", encoding="utf-8") as bat_file:
            bat_file.write(bat_content)

        subprocess.Popen(bat_path, shell=True, creationflags=subprocess.CREATE_NO_WINDOW)
        sys.exit()

    except Exception as e:
        try: progress_win.destroy() 
        except: pass
        messagebox.showerror("更新失敗", f"下載或替換檔案時發生錯誤：\n{e}")

def log_msg(msg):
    global execution_logs
    now = datetime.datetime.now()
    timestamp = now.strftime("%H:%M:%S.") + f"{now.microsecond // 1000:03d}"
    log_line = f"[{timestamp}] {msg}"
    print(log_line)
    execution_logs.append(log_line)

def clean_filename(text):
    return re.sub(r'[\\/*?:"<>|]', "", text).strip()

def get_executable_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))

os.chdir(get_executable_dir())

def force_quit_driver(driver_instance):
    try:
        subprocess.run("taskkill /F /IM msedge.exe /T", shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        subprocess.run("taskkill /F /IM msedgedriver.exe /T", shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except: pass

# --- 核心功能函數 ---

def get_driver():
    edge_options = Options()
    edge_options.page_load_strategy = 'eager' 
    edge_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    edge_options.add_experimental_option('useAutomationExtension', False)
    edge_options.add_argument("--disable-infobars")
    edge_options.add_argument("--disable-notifications")
    edge_options.add_argument("--disable-gpu")   
    edge_options.add_argument("--force-device-scale-factor=1")
    edge_options.add_argument("--disable-features=BlockInsecurePrivateNetworkRequests,PrivateNetworkAccessSendPreflights,IsolateOrigins,site-per-process")
    edge_options.add_argument("--disable-web-security")
    edge_options.add_argument("--allow-running-insecure-content")
    edge_options.add_argument("--allow-insecure-localhost")
    edge_options.add_argument("--ignore-certificate-errors")
    edge_options.add_argument("--remote-allow-origins=*")
    prefs = {
        "profile.default_content_setting_values.notifications": 2,
        "profile.managed_default_content_settings.insecure_private_network_requests": 1
    }
    edge_options.add_experimental_option("prefs", prefs)
    try:
        driver = webdriver.Edge(options=edge_options)
        try:
            import ctypes
            screen_w = ctypes.windll.user32.GetSystemMetrics(0)
            screen_h = ctypes.windll.user32.GetSystemMetrics(1)
            
            if screen_w > 1600:
                driver.set_window_position(0, 0)
                driver.set_window_size(1550, 1000)
            else:
                target_w = int(screen_w * 0.9)
                target_h = int(screen_h * 0.9)
                pos_x = int((screen_w - target_w) / 6)
                pos_y = int((screen_h - target_h) / 5)
                driver.set_window_position(pos_x, pos_y)
                driver.set_window_size(target_w, target_h)
        except:
            driver.maximize_window()

        try:
            driver.execute_cdp_cmd("Emulation.setDeviceMetricsOverride", {
                "width": 1550,             
                "height": 1000,            
                "deviceScaleFactor": 1,    
                "mobile": False
            })
        except Exception as cdp_e:
            pass 

        return driver
    except Exception as e:
        log_msg(f"瀏覽器啟動遇到問題: {e}")
        raise e
    
def logout():
    global driver
    try:
        if driver:
            log_msg("登出中，請稍候...")
            driver.get("https://stockservices.tdcc.com.tw/evote/logout.html")
            time.sleep(0.1) 
    except: pass

def close_tdcc_upload_tab_and_back(driver_instance, original_window=None, timeout=5):
    try:
        if not original_window:
            original_window = driver_instance.current_window_handle
    except:
        original_window = None

    end_time = time.time() + timeout
    handled = False

    while time.time() < end_time:
        try: handles = driver_instance.window_handles[:]
        except: break

        for handle in handles:
            if handle == original_window: continue
            try: driver_instance.switch_to.window(handle)
            except: continue

            try: current_url = driver_instance.current_url or ""
            except: current_url = ""

            if "/TDCCWEB/upload/" in current_url:
                try: driver_instance.close()
                except: pass
                handled = True
                break

        if handled: break
        time.sleep(0.2)

    try: handles = driver_instance.window_handles[:]
    except: handles = []

    try:
        if original_window and original_window in handles:
            driver_instance.switch_to.window(original_window)
        elif handles:
            driver_instance.switch_to.window(handles[0])
    except: pass    

def pass_active_form():
    global driver, join_draw
    try:
        form_=driver.find_element(By.ID, "msgDialog")
        if "抽獎" in form_.text:
            if join_draw:
                log_msg("🎉 偵測到抽獎視窗！依設定停留 5 分鐘 (300秒) 讓您手動參加...")
                time.sleep(300)
                log_msg("⏳ 5 分鐘結束，關閉提示繼續自動流程...")
            try:
                driver.find_element(By.ID, "msgDialog_okBtn").click()
            except:
                pass
    except:
        pass
        
    try:
        form_btn=driver.find_element(By.ID, "comfirmDialog_skipBtn")
        if "抽獎" in form_btn.text:
            if join_draw:
                log_msg("🎉 偵測到抽獎活動！依設定停留 5 分鐘 (300秒) 讓您手動參加...")
                time.sleep(300)
                log_msg("⏳ 5 分鐘結束，點擊略過繼續自動流程...")
            try:
                form_btn.click()
            except:
                pass
    except:
        pass

def autoLogin(user_ID):
    global driver, vote_speed, login_type
    log_msg(f"正在為您登入帳號: {user_ID}")
    
    base_wait = 0.1 * vote_speed
    try:
        driver.set_page_load_timeout(30) 
        driver.set_script_timeout(30)
        driver.implicitly_wait(0.2)
    except: pass

    try: driver.get("https://stockservices.tdcc.com.tw/evote/login/shareholder.html")
    except: pass

    input_timeout = max(2.0, 10.0 * vote_speed)
    start_wait = time.time()
    while time.time() - start_wait < input_timeout:

        try:
            msg_btns = driver.find_elements(By.ID, "msgDialog_okBtn")
            if msg_btns and msg_btns[0].is_displayed():
                log_msg("首頁偵測到系統對話框 (機器人驗證)，嘗試關閉...")
                msg_btns[0].click()
                time.sleep(0.1)
        except: pass
        
        try:
            robot_close = driver.find_elements(By.CSS_SELECTOR, 'button[onclick="$.modal.close();return false;"]')
            if robot_close and robot_close[0].is_displayed():
                log_msg("首頁偵測到機器人驗證/系統提示，嘗試關閉...")
                robot_close[0].click()
                time.sleep(0.1)
        except: pass

        try:
            driver.find_element(By.NAME,"pageIdNo").clear()
            driver.find_element(By.NAME,"pageIdNo").send_keys(user_ID)
            break
        except: time.sleep(base_wait)
    
    try: 
        log_msg(f"選擇登入方式: {login_type}")
        driver.find_element(By.NAME,"caType").send_keys(login_type)
    except: pass
    
    try: driver.find_element(By.ID, 'loginBtn').click()
    except: pass
    
    is_mobile_or_natural = False
    if login_type == "券商網路下單憑證":
        HARD_TIMEOUT_SECONDS = 20.0
        log_msg("等待券商憑證驗證 (限時20秒)...")
    else:
        is_mobile_or_natural = True
        HARD_TIMEOUT_SECONDS = 120.0
        log_msg(f"等待{login_type} (請注意手機/插卡，限時2分鐘)...")
    
    login_start_time = time.time()
    
    while True:
        if time.time() - login_start_time > HARD_TIMEOUT_SECONDS:
            log_msg(f"登入超時 (超過{int(HARD_TIMEOUT_SECONDS)}秒)，強制換下一個。")
            raise LoginTimeoutError("Timeout")

        time.sleep(base_wait*0.5) 

        try:
            msg_btns = driver.find_elements(By.ID, "msgDialog_okBtn")
            if msg_btns and msg_btns[0].is_displayed():
                log_msg("登入期間偵測到系統對話框 (機器人驗證)，嘗試關閉...")
                msg_btns[0].click()
                time.sleep(0.1)
                try: 
                    driver.find_element(By.ID, 'loginBtn').click() # 關閉後立刻再按一次登入
                    log_msg("已重新點擊登入按鈕")
                except: pass
        except: pass
        
        try:
            robot_close = driver.find_elements(By.CSS_SELECTOR, 'button[onclick="$.modal.close();return false;"]')
            if robot_close and robot_close[0].is_displayed():
                log_msg("登入等待期間偵測到機器人驗證/系統提示，嘗試關閉...")
                robot_close[0].click()
                time.sleep(0.1)
                try: 
                    driver.find_element(By.ID, 'loginBtn').click() # 關閉後立刻再按一次登入
                    log_msg("已重新點擊登入按鈕")
                except: pass
        except: pass
        
                
        try:
            from selenium.webdriver.support import expected_conditions as EC
            from selenium.webdriver.support.ui import WebDriverWait
            from selenium.common.exceptions import TimeoutException
            
            try:
                WebDriverWait(driver, 1).until(EC.alert_is_present())
                alert = driver.switch_to.alert
                log_msg(f"偵測到網頁訊息，嘗試關閉...")
                alert.accept() 
                raise LoginTimeoutError("Cert Error Alert")
            except TimeoutException:
                pass
        except LoginTimeoutError: raise
        except: pass

        try:
            current_url = driver.current_url
            login_success = False
            
            if is_mobile_or_natural:
                if "tc_estock_welshas" in current_url:
                    login_success = True
            else:
                if "login/shareholder" not in current_url:
                    login_success = True

            if login_success:
                if "tc_estock_welshas" not in current_url:
                     driver.get("https://stockservices.tdcc.com.tw/evote/shareholder/000/tc_estock_welshas.html")
                
                try: driver.execute_script("document.body.style.zoom = '100%'")
                except: pass
                break
        except: pass

        try:
            if "系統維護中" in driver.find_element(By.TAG_NAME,'body').text:
                raise SystemMaintenanceError("System Maintenance")
        except SystemMaintenanceError: raise
        except: pass

        # ------------------- 登入後的雜項彈窗清理 (包含Email抽獎、同意條款等) -------------------
        for _ in range(3): # 循環3次，對付連續跳出的多重彈窗
            # 1. 通用系統確認按鈕
            try: 
                btn = driver.find_element(By.ID, "comfirmDialog_okBtn")
                if btn.is_displayed(): btn.click(); time.sleep(0.5)
            except: pass
            
            # 2. 抽獎/Email 專用略過按鈕
            try:
                skip_btn = driver.find_element(By.ID, "comfirmDialog_skipBtn")
                if skip_btn.is_displayed(): skip_btn.click(); time.sleep(0.5)
            except: pass
            
            # 3. 找尋「略過」、「稍後」、「不參加」等文字按鈕 (對付各種 Email 填寫變化)
            try:
                text_btns = driver.find_elements(By.XPATH, "//button[contains(text(),'略過')] | //button[contains(text(),'稍後')] | //button[contains(text(),'不參加')] | //a[contains(text(),'略過')] | //a[contains(text(),'稍後')]")
                for tb in text_btns:
                    if tb.is_displayed():
                        tb.click()
                        time.sleep(0.5)
                        break
            except: pass

            # 4. 同意條款流程 (閱讀條款 -> 打勾 -> 同意)
            try:
                agree_link = driver.find_element(By.CSS_SELECTOR, 'a[id="agreeLink"]')
                if agree_link.is_displayed():
                    original_window = driver.current_window_handle
                    old_handles = set(driver.window_handles)
                    agree_link.click()
                    for _ in range(20):
                        try:
                            if len(set(driver.window_handles) - old_handles) > 0: break
                        except: pass
                        time.sleep(0.2)
                    close_tdcc_upload_tab_and_back(driver, original_window=original_window, timeout=5)
                    time.sleep(0.5)
            except: pass

            try:
                agree_terms = driver.find_element(By.CSS_SELECTOR, 'input[id="agreeTerms"]')
                if agree_terms.is_displayed() and not agree_terms.is_selected():
                    agree_terms.click(); time.sleep(0.5)
            except: pass

            try:
                agree_btn = driver.find_element(By.CSS_SELECTOR, 'a[class="btnAgree btn-style btn-b btn-lg"]')
                if agree_btn.is_displayed(): agree_btn.click(); time.sleep(0.5)
            except: pass

            # 5. 其他莫名其妙的 btn1
            try:
                btn1 = driver.find_element(By.NAME, 'btn1')
                if btn1.is_displayed(): btn1.click(); time.sleep(0.5)
            except: pass

def process_single_revoke():
    global driver, vote_speed
    base_wait = 0.5 * vote_speed
    try:
        log_msg("進入撤銷確認頁面...")
        time.sleep(base_wait)
        
        for step in range(1, 4): 
            confirm_clicked = False
            log_msg(f"嘗試尋找第 {step} 次確認按鈕或提示框...")
            time.sleep(base_wait)
            
            btns = driver.find_elements(By.XPATH, "//button[contains(text(),'確認')] | //a[contains(text(),'確認')] | //button[contains(text(),'確定')] | //input[@value='確認' or @value='確定']")
            for btn in btns:
                try:
                    if btn.is_displayed():
                        btn.click()
                        confirm_clicked = True
                        log_msg(f"已點擊第 {step} 次確認 (網頁按鈕)。")
                        break
                except: pass
                
            if not confirm_clicked:
                try:
                    msg_btn = driver.find_element(By.ID, "msgDialog_okBtn")
                    if msg_btn.is_displayed():
                        msg_btn.click()
                        confirm_clicked = True
                        log_msg(f"已同意第 {step} 次確認 (系統提示視窗)。")
                except: pass

            if not confirm_clicked:
                try:
                    alert = driver.switch_to.alert
                    alert_text = alert.text
                    alert.accept()
                    confirm_clicked = True
                    log_msg(f"已同意第 {step} 次確認 (瀏覽器警告)，內容: [{alert_text}]")
                except: pass

        log_msg("等待手動介入選擇憑證... (請在跳出的視窗中操作)")
        log_msg("程式正在背景偵測彈出的視窗，並會自動幫您點擊最後的確認按鈕...")
        
        timeout_wait = 300 
        start_wait = time.time()
        
        initial_windows = driver.window_handles
        logged_windows = set(initial_windows)
        logged_modals = set()

        while time.time() - start_wait < timeout_wait:
            try:
                current_windows = driver.window_handles
                for w in current_windows:
                    if w not in logged_windows:
                        driver.switch_to.window(w)
                        w_title = driver.title
                        w_url = driver.current_url
                        w_text = driver.find_element(By.TAG_NAME, "body").text.replace('\n', ' ')[:100] 
                        log_msg(f"[偵測到新視窗] 標題: {w_title}")
                        log_msg(f"   - 網址: {w_url}")
                        log_msg(f"   - 內容預覽: {w_text}...")
                        logged_windows.add(w)
                        driver.switch_to.window(initial_windows[0]) 
            except Exception:
                try: driver.switch_to.window(initial_windows[0])
                except: pass

            try:
                dialogs = driver.find_elements(By.CSS_SELECTOR, "div.modal, div.dialog, div[role='dialog'], iframe")
                for d in dialogs:
                    if d.is_displayed():
                        d_text = d.text.strip().replace('\n', ' ')
                        if d_text and d_text not in logged_modals and len(d_text) > 5:
                            log_msg(f"[偵測到網頁對話框] 內容: {d_text[:100]}...")
                            logged_modals.add(d_text)
            except: pass

            try:
                final_msg_btn = driver.find_element(By.ID, "msgDialog_okBtn")
                if final_msg_btn.is_displayed():
                    final_msg_btn.click()
                    log_msg("已自動點擊憑證簽署後的最後確認按鈕 (msgDialog)！")
                    time.sleep(1) 
            except: pass

            try:
                final_alert = driver.switch_to.alert
                alert_msg = final_alert.text
                final_alert.accept()
                log_msg(f"已自動同意憑證簽署後的警告視窗: [{alert_msg}]")
                time.sleep(1)
            except: pass

            try:
                final_btns = driver.find_elements(By.XPATH, "//button[contains(text(),'確認')] | //a[contains(text(),'確認')] | //button[contains(text(),'確定')] | //input[@value='確認' or @value='確定']")
                for f_btn in final_btns:
                    if f_btn.is_displayed():
                        f_btn.click()
                        log_msg("已自動點擊憑證簽署後的最後確認按鈕 (網頁元素)！")
                        time.sleep(1)
                        break
            except: pass

            try:
                if "tc_estock_welshas" in driver.current_url and driver.find_elements(By.NAME, 'qryStockId'):
                    log_msg("偵測到已返回列表頁面，代表撤銷流程完整結束，接續下一筆。")
                    time.sleep(1) 
                    return True
            except: pass
            
            time.sleep(1)
            
        log_msg("等待手動操作與最後確認超時！")
        return False
        
    except Exception as e:
        log_msg(f"撤銷流程發生錯誤: {e}")
        return False

def auto_revoke(user_id, mode, stock_list):
    global driver
    base_wait = 0.5
    try:
        if "tc_estock_welshas" not in driver.current_url:
            driver.get("https://stockservices.tdcc.com.tw/evote/shareholder/000/tc_estock_welshas.html")
        
        for _ in range(50): 
            if driver.find_elements(By.NAME,'qryStockId'): break
            time.sleep(0.1)

        if mode == "specific":
            for stock_id in stock_list:
                log_msg(f"正在搜尋並嘗試撤銷代號: {stock_id}")
                driver.find_element(By.NAME,'qryStockId').clear()
                driver.find_element(By.NAME,'qryStockId').send_keys(stock_id)
                driver.find_element(By.CSS_SELECTOR,'a[onclick="qryByStockId();"]').click()
                time.sleep(1)
                
                try:
                    rows = driver.find_elements(By.TAG_NAME,'tr')
                    if len(rows) > 1 and str(stock_id) in rows[1].text:
                        revoke_links = rows[1].find_elements(By.XPATH, ".//a[contains(text(),'撤銷')]")
                        if revoke_links:
                            revoke_links[0].click()
                            process_single_revoke()
                        else:
                            log_msg(f"[{stock_id}] 找不到撤銷按鈕，可能尚未投票或無法撤銷。")
                    else:
                        log_msg(f"找不到代號: {stock_id}")
                except Exception as e:
                    log_msg(f"搜尋或點擊 [{stock_id}] 時出錯: {e}")
                
                if "tc_estock_welshas" not in driver.current_url:
                     driver.get("https://stockservices.tdcc.com.tw/evote/shareholder/000/tc_estock_welshas.html")
                     time.sleep(1)

        elif mode == "all":
            log_msg("開始掃描並撤銷所有已投票項目...")
            while True:
                found_revoke = False
                try:
                    trs = driver.find_elements(By.TAG_NAME,'tr')
                    if len(trs) > 1:
                        for row in trs[1:]:
                            try:
                                revoke_links = row.find_elements(By.XPATH, ".//a[contains(text(),'撤銷')]")
                                if revoke_links and revoke_links[0].is_displayed():
                                    row_text = row.text.split(" ")
                                    t_id = row_text[0]
                                    log_msg(f"發現可撤銷項目: {t_id}")
                                    revoke_links[0].click()
                                    found_revoke = True
                                    process_single_revoke()
                                    break 
                            except StaleElementReferenceException:
                                continue
                except Exception as e:
                    log_msg(f"掃描列表錯誤: {e}")
                
                if not found_revoke:
                    log_msg("------ 本帳號已無可撤銷項目 ------")
                    break

        return 0
    except Exception as e:
        log_msg(f"撤銷任務失敗: {e}")
        return 1

def voting():
    global default_vote, manual_vote, accept_list, opposite_list, abstain_list, driver, vote_speed
    try: driver.implicitly_wait(0)
    except: pass
    base_wait = 0.1 * vote_speed
    start_time = time.time()
    max_vote_time = max(30.0, 300.0 * vote_speed)

    log_msg("進入投票頁面，開始偵測按鈕與流程...")

    # 用來避免某些重複洗頻的 Log
    logged_hidden_robot = False

    while time.time() - start_time < max_vote_time: 
        try:
            time.sleep(base_wait) 
            
            # 1. 處理機器人驗證與系統提示框 ($.modal.close)
            try:
                robot_close = driver.find_elements(By.CSS_SELECTOR, 'button[onclick="$.modal.close();return false;"]')
                if robot_close:
                    if robot_close[0].is_displayed():
                        log_msg("偵測到『機器人驗證或系統提示』視窗！嘗試點擊關閉...")
                        robot_close[0].click()
                        log_msg("已發送關閉指令。")
                        time.sleep(base_wait * 3)
                        continue
                    elif not logged_hidden_robot:
                        log_msg("發現機器人驗證按鈕(可能被遮擋或尚未載入完成)")
                        logged_hidden_robot = True # 只印一次避免洗畫面
            except Exception as e:
                log_msg(f"處理機器人驗證時發生錯誤: {e}")

            # 2. 處理「略過」按鈕
            try:
                ignore_btns = driver.find_elements(By.CSS_SELECTOR, 'button[onclick*="ignoreVote"]')
                if ignore_btns and ignore_btns[0].is_displayed():
                    log_msg("偵測到『略過』按鈕，嘗試點擊...")
                    ignore_btns[0].click()
                    log_msg("已點擊略過。")
                    time.sleep(base_wait * 5)
                    continue
            except Exception as e:
                pass # 避免干擾主流程

            # 3. 處理「選擇權未行使」等特殊頁面
            try:
                body_list = driver.find_elements(By.TAG_NAME, "body")
                if body_list:
                    body_text = body_list[0].text
                    if "選擇權" in body_text and "未行使" in body_text:
                        clicked_backup = False
                        code_btns = driver.find_elements(By.CSS_SELECTOR, "button[onclick*='Next'], button[onclick*='goNext'], a[onclick*='Next']")
                        for btn in code_btns:
                            if btn.is_displayed():
                                log_msg("處理選擇權頁面，點擊下一步 (代碼綁定)...")
                                btn.click(); clicked_backup = True; break
                        
                        if not clicked_backup:
                            text_btns = driver.find_elements(By.XPATH, "//button[contains(text(),'下一步')] | //a[contains(text(),'下一步')]")
                            if text_btns and text_btns[0].is_displayed():
                                log_msg("處理選擇權頁面，點擊下一步 (文字綁定)...")
                                text_btns[0].click(); clicked_backup = True
                        
                        if clicked_backup:
                            time.sleep(base_wait * 2)
                            continue
            except StaleElementReferenceException: continue 
            except: pass

            # 4. 處理一般訊息對話框 (msgDialog)
            try:
                msg_btns = driver.find_elements(By.ID, "msgDialog_okBtn")
                if msg_btns and msg_btns[0].is_displayed():
                    log_msg("偵測到系統訊息對話框，嘗試點擊確認...")
                    msg_btns[0].click()
                    continue
            except: pass

            # 5. 處理完成按鈕 (doProcess) -> 這是成功投完的最後一步
            try:
                finish_btns = driver.find_elements(By.CSS_SELECTOR, 'button[onclick="doProcess();"]')
                if finish_btns and finish_btns[0].is_displayed():
                    log_msg("偵測到『完成投票』按鈕，點擊送出！")
                    finish_btns[0].click()
                    return True # 成功結束
            except: pass

            # 6. 執行投票動作 (點擊贊成/反對/棄權)
            clicked_vote = False
            try:
                js_script = ""
                if default_vote == "accept":
                    js_script = "var a = document.querySelector('table.c-votelist_docSection tr:nth-child(2) td:nth-child(2) a:nth-child(1)'); if(a) { a.click(); return true; } else { return false; }"
                elif default_vote == "opposite":
                    js_script = "var a = document.querySelector('table.c-votelist_docSection tr:nth-child(2) td:nth-child(2) a:nth-child(2)'); if(a) { a.click(); return true; } else { return false; }"
                elif default_vote == "abstain":
                    js_script = "var a = document.querySelector('table.c-votelist_docSection tr:nth-child(2) td:nth-child(2) a:nth-child(3)'); if(a) { a.click(); return true; } else { return false; }"
                
                if js_script:
                    if driver.execute_script(js_script): 
                        if not getattr(driver, '_logged_vote_action', False): # 簡單防止重複洗頻
                            log_msg(f"已透過腳本執行預設投票: {default_vote}")
                            driver._logged_vote_action = True
                        clicked_vote = True

                if manual_vote:
                     rows = driver.find_elements(By.XPATH,'//td/input[@type="radio"]/../..')
                     for row in rows:
                         txt = row.text
                         if any(k in txt for k in accept_list):
                             r_btns = row.find_elements(By.CSS_SELECTOR, 'input[value="A"]')
                             if r_btns: r_btns[0].click(); clicked_vote = True
                         elif any(k in txt for k in opposite_list):
                             r_btns = row.find_elements(By.CSS_SELECTOR, 'input[value="O"]')
                             if r_btns: r_btns[0].click(); clicked_vote = True
                         elif any(k in txt for k in abstain_list):
                             r_btns = row.find_elements(By.CSS_SELECTOR, 'input[value="C"]')
                             if r_btns: r_btns[0].click(); clicked_vote = True
                     if clicked_vote:
                         log_msg("已依據關鍵字清單執行手動投票。")
            except: pass

            if clicked_vote: time.sleep(base_wait)

            # 7. 處理推進流程的按鈕 (確認投票結果、下一步)
            action_taken = False
            
            # 7-1. 確認投票結果
            confirm_btns = driver.find_elements(By.XPATH, "//button[contains(text(),'確認投票結果')] | //a[contains(text(),'確認投票結果')]")
            for btn in confirm_btns:
                try:
                    if btn.is_displayed():
                        log_msg("準備點擊『確認投票結果』...")
                        btn.click()
                        action_taken = True
                        break
                except StaleElementReferenceException: continue 

            # 7-2. 下一步
            if not action_taken:
                next_btns = driver.find_elements(By.XPATH, "//button[contains(text(),'下一步')] | //a[contains(text(),'下一步')]")
                for btn in next_btns:
                    try:
                        if btn.is_displayed():
                            log_msg("準備點擊『下一步』...")
                            btn.click()
                            action_taken = True
                            time.sleep(base_wait * 2)
                            break 
                    except StaleElementReferenceException: continue

            # 7-3. 其他腳本觸發的流程按鈕
            if not action_taken:
                selectors = ['button[onclick="voteObj.checkVote(); return false;"]', 'button[onclick="voteObj.ignoreVote();voteObj.goNext(); return false;"]', 'button[onclick="voteObj.checkMeetingPartner(); return false;"]']
                for sel in selectors:
                    btns = driver.find_elements(By.CSS_SELECTOR, sel)
                    for btn in btns:
                        try:
                            if btn.is_displayed():
                                log_msg("點擊系統流程輔助按鈕...")
                                btn.click()
                                action_taken = True
                                break 
                        except StaleElementReferenceException: continue
                    if action_taken: break
            
            # 8. 處理放棄按鈕 (防呆)
            if not action_taken:
                gu_btn = driver.find_elements(By.CSS_SELECTOR, 'a[href="javascript:giveUp();"]')
                if gu_btn and gu_btn[0].is_displayed():
                    log_msg("發現『放棄』按鈕，執行點擊...")
                    gu_btn[0].click()

        except StaleElementReferenceException:
            # 這是 Selenium 常見狀況，代表網頁正在刷新，元素失效，直接繼續下一次迴圈即可
            continue
        except Exception as e:
            # 捕捉其他未預期的錯誤，但不中斷迴圈
            pass

    log_msg("投票流程超時，未能完成該筆投票。")
    return False

def autovote(user_ID):
    global driver, voteinfolist, session_results, vote_speed
    try: driver.implicitly_wait(0)
    except: pass
    
    if user_ID not in session_results:
        session_results[user_ID] = {'success': [], 'fail_vote': [], 'fail_screenshot': [], 'success_screenshot': []}

    failed_attempts = set()
    base_wait = 0.1 * vote_speed

    try:
        log_msg("開始自動掃描未投票項目...")
        pass_active_form()
        if "tc_estock_welshas" not in driver.current_url:
            driver.get("https://stockservices.tdcc.com.tw/evote/shareholder/000/tc_estock_welshas.html")
        
        for _ in range(50): 
            if driver.find_elements(By.NAME,'qryStockId'): break
            time.sleep(base_wait)
            
        log_msg("------ 列表掃描中 ------")
        while True:
            pass_active_form()
            try:
                trs = driver.find_elements(By.TAG_NAME,'tr')
                target_row = None; target_stock_id = ""; target_stock_name = ""

                if len(trs) > 1:
                    for row in trs[1:]:
                        try:
                            row_text = row.text
                            if "未投票" in row_text:
                                parts = row_text.split(" ")
                                s_id = parts[0]
                                if s_id in failed_attempts: continue
                                target_row = row
                                target_stock_id = s_id
                                target_stock_name = parts[1] if len(parts) > 1 else ""
                                break
                        except StaleElementReferenceException:
                             continue 

                if target_row:
                    log_msg(f"發現未投票: {target_stock_id} {target_stock_name}")
                    try:
                        tds = target_row.find_elements(By.TAG_NAME,'td')
                        if len(tds) > 3:
                            btns = tds[3].find_elements(By.TAG_NAME,'a')
                            if btns:
                                try:
                                    btns[0].click()
                                except StaleElementReferenceException:
                                    log_msg(f"[{target_stock_id}] 點擊瞬間頁面刷新，嘗試繼續...")
                                    pass
                                
                                time.sleep(base_wait * 2) 
                                
                                msg_btns = driver.find_elements(By.ID, "msgDialog_okBtn")
                                if msg_btns and msg_btns[0].is_displayed():
                                    try: msg_btns[0].click()
                                    except: pass

                                is_success = voting()
                                record_text = f"{target_stock_id} {target_stock_name}".strip()
                                
                                if is_success:
                                    if user_ID not in voteinfolist: voteinfolist[user_ID] = {}
                                    if isinstance(voteinfolist[user_ID], list): 
                                        voteinfolist[user_ID] = {sid: 0 for sid in voteinfolist[user_ID]}
                                    
                                    if target_stock_id not in voteinfolist[user_ID]:
                                        voteinfolist[user_ID][target_stock_id] = 0
                                    write_voteinfolist(voteinfolist)
                                    session_results[user_ID]['success'].append(record_text)
                                    log_msg(f"[{target_stock_id}] 投票成功！")
                                else:
                                    session_results[user_ID]['fail_vote'].append(record_text)
                                    log_msg(f"[{target_stock_id}] 投票未完成")
                                    failed_attempts.add(target_stock_id)
                                
                                if "tc_estock_welshas" not in driver.current_url:
                                     driver.get("https://stockservices.tdcc.com.tw/evote/shareholder/000/tc_estock_welshas.html")
                            else: failed_attempts.add(target_stock_id)
                        else: failed_attempts.add(target_stock_id)
                    except Exception as e:
                        log_msg(f"單筆處理異常 ({target_stock_id}): {e}")
                        failed_attempts.add(target_stock_id)
                        driver.get("https://stockservices.tdcc.com.tw/evote/shareholder/000/tc_estock_welshas.html")
                else:
                    log_msg(f"------ 本帳號已無待辦事項 ------")
                    break
            except Exception as e:
                log_msg(f"列表掃描錯誤: {e}"); return 1
        return 0
    except Exception as e:
        log_msg(f"Autovote 執行失敗: {e}"); return 1

def screenshot(user_id, info):
    global base_path, screenshot_mode, driver, user_name_map, shot_speed # <--- 修正：將 time_speed 改為 shot_speed
    try:
        driver.execute_script("""
            document.body.style.width = '1600px';      
            document.body.style.minWidth = '1600px';   
            document.body.style.marginLeft = '0px';
            document.body.style.zoom = '1.1'; 
        """)
        time.sleep(0.2 * shot_speed) # <--- 修正：套用分離後的截圖等待速度

        x_start, y_start = 255, 170
        width_rect, height_rect = 1248, 339
        
        body_el = driver.find_element(By.TAG_NAME, "body")
        png_data = body_el.screenshot_as_png
        
        img = Image.open(io.BytesIO(png_data))
        
        left = x_start
        top = y_start
        right = x_start + width_rect
        bottom = y_start + height_rect
        
        cropped_img = img.crop((left, top, min(img.width, right), min(img.height, bottom)))
        
        display_name = user_name_map.get(user_id, str(user_id))
        stock_name = info[1].replace('*','')
        stock_name = clean_filename(stock_name)
        if len(stock_name) > 20: stock_name = stock_name[:20]

        date_prefix = datetime.datetime.now().strftime("%Y%m%d")

        if screenshot_mode == 1:
            save_dir = os.path.join(base_path, display_name)
            filename = f"{date_prefix}_{info[0]}_{stock_name}.png"
        else:
            safe_display_name = display_name if len(display_name) < 10 else display_name[:10]
            save_dir = base_path
            filename = f"{date_prefix}_{info[0]}_{stock_name}_{safe_display_name}.png"

        if not os.path.exists(save_dir): os.makedirs(save_dir)
        cropped_img.save(os.path.join(save_dir, filename))
        
        try: img.close()
        except: pass
        
        log_msg(f"截圖已保存: {filename}")
        return 0
    except Exception as e:
        log_msg(f"截圖失敗: {e}")
        return 1

def auto_screenshot(user_id, stock_id):
    # 此處已刪除重複出現的函式，並確保只保留這一個正確的修正版
    global driver, session_results, user_name_map, shot_speed
    try: driver.implicitly_wait(0.01)
    except: pass
    
    if user_id not in session_results:
        session_results[user_id] = {'success': [], 'fail_vote': [], 'fail_screenshot': [], 'success_screenshot': []}
    if 'success_screenshot' not in session_results[user_id]:
        session_results[user_id]['success_screenshot'] = []

    base_wait = 0.1 * shot_speed
    try:
        log_msg(f"搜尋股票: {stock_id}")
        pass_active_form()
        if "tc_estock_welshas.html" not in driver.current_url:
            driver.get("https://stockservices.tdcc.com.tw/evote/shareholder/000/tc_estock_welshas.html")
        
        try: driver.execute_script("document.body.style.zoom = '100%'")
        except: pass

        for _ in range(100):
            try:
                driver.find_element(By.NAME,'qryStockId')
                break
            except: time.sleep(base_wait)
            
        driver.find_element(By.NAME,'qryStockId').clear()
        driver.find_element(By.NAME,'qryStockId').send_keys(stock_id)
        driver.find_element(By.CSS_SELECTOR,'a[onclick="qryByStockId();"]').click()
        
        found_result = False
        for _ in range(100): 
            time.sleep(base_wait)
            try:
                rows = driver.find_elements(By.TAG_NAME,'tr')
                if len(rows) > 1 and str(stock_id) in rows[1].text:
                    found_result = True
                    break
            except: pass
            
        if not found_result:
            log_msg(f"找不到代號: {stock_id}")
            session_results[user_id]['fail_screenshot'].append(stock_id)
            return 2

        voteinfo = []
        try:
            row = driver.find_elements(By.TAG_NAME,'tr')[1]
            parts = row.text.split(" ")
            voteinfo.extend(parts[0:2])
            report_text = f"{voteinfo[0]} {voteinfo[1]}".strip() if len(voteinfo) > 1 else stock_id

            page_loaded = False
            for attempt in range(5):
                try:
                    if driver.find_elements(By.CSS_SELECTOR, 'button[onclick*="back"], input[onclick*="back"]'):
                        page_loaded = True
                        break
                    if driver.find_elements(By.ID, "msgDialog_okBtn") and driver.find_element(By.ID, "msgDialog_okBtn").is_displayed():
                        page_loaded = True
                        break
                except: pass

                try:
                    current_row = driver.find_elements(By.TAG_NAME,'tr')[1]
                    links = current_row.find_elements(By.TAG_NAME,'a')
                    target_link = None
                    for link in links:
                        if "查詢" in link.text:
                            target_link = link
                            break
                    if target_link:
                        if attempt > 0: log_msg(f"第 {attempt+1} 次嘗試進入頁面...")
                        target_link.click()
                        check_limit = 100 
                        for _ in range(check_limit):
                            if driver.find_elements(By.CSS_SELECTOR, 'button[onclick*="back"], input[onclick*="back"]'):
                                page_loaded = True
                                break
                            if driver.find_elements(By.ID, "msgDialog_okBtn") and driver.find_element(By.ID, "msgDialog_okBtn").is_displayed():
                                page_loaded = True
                                break
                            time.sleep(base_wait)
                        if page_loaded: break 
                    else:
                        time.sleep(base_wait * 5)
                        continue
                except: time.sleep(base_wait * 2)
            
            if not page_loaded:
                try:
                    nav_btns = driver.find_elements(By.XPATH, "//button | //a | //input[@type='button']")
                    for btn in nav_btns:
                        if not btn.is_displayed(): continue
                        txt = btn.text.strip()
                        val = btn.get_attribute("value")
                        check_str = (txt + str(val)).strip()
                        if "返回" in check_str or "上一頁" in check_str or "列表" in check_str:
                            page_loaded = True
                            break
                except: pass

            if not page_loaded:
                log_msg(f"[{stock_id}] 進入內頁失敗，跳過截圖")
                session_results[user_id]['fail_screenshot'].append(report_text)
                return 2

            try: driver.execute_script("document.body.style.zoom = '100%'")
            except: pass

            if user_id in user_name_map:
                for _ in range(10):
                    try:
                        btn = driver.find_element(By.ID, "msgDialog_okBtn")
                        if btn.is_displayed():
                            btn.click()
                            break 
                    except: pass
                    time.sleep(base_wait)
            else:
                detected_name = ""
                start_search = time.time()
                while time.time() - start_search < 2.0:
                    try:
                        btn = driver.find_element(By.ID, "msgDialog_okBtn")
                        if btn.is_displayed(): btn.click()
                    except: pass
                    try: driver.execute_script("document.body.style.zoom = '100%'")
                    except: pass
                    try:
                        xpath_targets = driver.find_elements(By.XPATH, "//*[contains(text(),'戶名')]")
                        for el in xpath_targets:
                            if not el.is_displayed(): continue
                            raw_text = el.text.strip()
                            if (":" in raw_text or "：" in raw_text) and len(raw_text) > 4:
                                detected_name = raw_text.replace("戶名", "").replace(":", "").replace("：", "").strip()
                                break
                            try:
                                parent = el.find_element(By.XPATH, "./..")
                                if parent.tag_name.lower() in ['tr', 'td', 'div']:
                                    full_text = parent.text.replace("\n", " ")
                                    if "戶名" in full_text:
                                        parts_t = full_text.split("戶名")
                                        if len(parts_t) > 1:
                                            potential = parts_t[1].strip(":： ").split(" ")[0]
                                            if len(potential) > 1:
                                                detected_name = potential
                                                break
                            except: pass
                        if detected_name:
                            clean_name = clean_filename(detected_name)
                            user_name_map[user_id] = clean_name
                            log_msg(f"抓到戶名: {clean_name}")
                            break
                    except: pass
                    time.sleep(base_wait)

            voteinfo.append("unknown") 
            res = screenshot(user_id, voteinfo)
            if res != 0: session_results[user_id]['fail_screenshot'].append(report_text)
            else: session_results[user_id]['success_screenshot'].append(report_text)
            
            try: driver.execute_script("arguments[0].click();", driver.find_element(By.CSS_SELECTOR,'button[onclick="back(); return false;"]'))
            except: pass
            return res
        except Exception as e:
            log_msg(f"流程錯誤: {e}")
            fail_record = locals().get('report_text', stock_id)
            session_results[user_id]['fail_screenshot'].append(fail_record)
            return 1
    except: return 1

def write_voteinfolist(voteinfolist):
    base = os.path.join(CONFIG_DIR, "queue_data")
    if not os.path.exists(base): os.makedirs(base)
    for id, stock_data in voteinfolist.items():
        safe_dirname = get_anonymous_dirname(id)
        user_path = os.path.join(base, safe_dirname)
        if not os.path.exists(user_path): os.makedirs(user_path)
        with open(os.path.join(user_path, "pending_queue.txt"), 'w', encoding='utf-8') as f:
            if isinstance(stock_data, list):
                f.writelines([f"{sid},0\n" for sid in stock_data])
            else:
                f.writelines([f"{sid},{count}\n" for sid, count in stock_data.items()])

def read_voteinfolist(voteinfolist):
    base = os.path.join(CONFIG_DIR, "queue_data")
    if not os.path.exists(base): return
    global shareholderIDs
    for uid in shareholderIDs:
        if not uid: continue
        safe_dirname = get_anonymous_dirname(uid)
        user_path = os.path.join(base, safe_dirname)
        file_path = os.path.join(user_path, "pending_queue.txt")
        if os.path.exists(file_path):
            with open(file_path, 'r', encoding='utf-8') as f:
                stocks_dict = {}
                for line in f.readlines():
                    line = line.strip()
                    if not line: continue
                    if ',' in line:
                        parts = line.split(',')
                        stocks_dict[parts[0]] = int(parts[1])
                    else:
                        stocks_dict[line] = 0
                if stocks_dict: voteinfolist[uid] = stocks_dict

def generate_session_report(start_t=None, end_t=None, count=0):
    global execution_logs, vote_speed, shot_speed, login_type # 引入全域變數
    try:
        log_dir = "Log"
        if not os.path.exists(log_dir):
            os.makedirs(log_dir)
            
        file_name = f"自動化任務報告_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
        report_name = os.path.join(log_dir, file_name)
        
        def chunks(lst, n):
            for i in range(0, len(lst), n):
                yield lst[i:i + n]

        with open(report_name, "w", encoding="utf-8") as f:
            f.write("=== 股東e票通 - 自動化任務報告 ===\n")
            f.write(f"時間: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            
            # --- 新增：速率與環境紀錄 ---
            f.write(f"【執行環境設定】\n")
            f.write(f"  - 登入方式: {login_type}\n")
            f.write(f"  - 投票操作速率: {vote_speed}\n")
            f.write(f"  - 截圖等待速率: {shot_speed}\n")
            # --------------------------

            if start_t and end_t:
                total_seconds = end_t - start_t
                avg_seconds = total_seconds / count if count > 0 else 0
                f.write(f"總耗時: {total_seconds:.2f} 秒\n")
                f.write(f"處理家數: {count}\n")
                f.write(f"平均每間耗時: {avg_seconds:.2f} 秒\n")
            f.write("\n")
            
            for uid, res in session_results.items():
                d_name = user_name_map.get(uid, uid)
                f.write(f"【帳號: {d_name} ({uid})】\n")
                f.write(f"  - 成功投票/撤銷: {len(res.get('success', []))}\n")
                if res.get('success', []): 
                    for chunk in chunks(res['success'], 5):
                        f.write(f"    內容: {', '.join(chunk)}\n")
                succ_shots = res.get('success_screenshot', [])
                f.write(f"  - 截圖成功: {len(succ_shots)}\n")
                if succ_shots:
                    for chunk in chunks(succ_shots, 5):
                        f.write(f"    內容: {', '.join(chunk)}\n")
                f.write(f"  - 截圖失敗: {len(res.get('fail_screenshot', []))}\n")
                if res.get('fail_screenshot', []): 
                    for chunk in chunks(res['fail_screenshot'], 5):
                        f.write(f"    內容: {', '.join(chunk)}\n")
                f.write(f"  - 失敗項目: {len(res.get('fail_vote', []))}\n")
                if res.get('fail_vote', []): 
                    for chunk in chunks(res['fail_vote'], 5):
                        f.write(f"    內容: {', '.join(chunk)}\n")
                f.write("-" * 30 + "\n")
            
            f.write("\n\n==========================================\n")
            f.write("           詳細執行歷程          \n")
            f.write("==========================================\n")
            for line in execution_logs:
                f.write(line + "\n")
                
        execution_logs.clear()
        
        log_msg(f"報告與詳細歷程已產生: {report_name}")
    except Exception as e:
        log_msg(f"報告產生失敗: {e}")

class TextRedirector(object):
    def __init__(self, widget, tag="stdout"):
        self.widget = widget
        self.tag = tag
    def write(self, str):
        self.widget.configure(state="normal")
        self.widget.insert("end", str, (self.tag,))
        self.widget.see("end")
        self.widget.configure(state="disabled")
    def flush(self): pass

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("股東e票通 - 輕鬆投自動化助手")
        
        # 預先載入設定，以取得儲存的視窗大小
        self.load_config()
        
        # 套用記憶的主視窗大小，或使用預設大小
        if main_window_geom:
            self.geometry(main_window_geom)
        else:
            window_width = 615
            window_height = 680
            screen_width = self.winfo_screenwidth()
            screen_height = self.winfo_screenheight()      
            x_cordinate = int(screen_width / 8) 
            y_cordinate = int(screen_height / 10)
            self.geometry(f"{window_width}x{window_height}+{x_cordinate}+{y_cordinate}")
        
        # 確保視窗可以調整大小，但不要縮到太小影響版面
        self.minsize(500, 500)
        self.resizable(True, True)
        
        # 攔截關閉事件，儲存視窗大小再關閉
        self.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        self.style = ttk.Style(self)
        try: self.style.theme_use('clam') 
        except: pass
        
        main_font = ('Microsoft JhengHei', 10)
        bold_font = ('Microsoft JhengHei', 10, 'bold')
        # --- UI 極簡黑白專業風 (搭配重點紅) ---
        app_bg = '#E8E8E8'     # 改為灰色背景
        text_color = '#111111' # 深黑文字
        
        self.configure(bg=app_bg) 

        self.style.configure('.', font=main_font, background=app_bg, foreground=text_color)
        self.style.configure('TFrame', background=app_bg)
        
        # 標題外框：黑白對比
        self.style.configure('TLabelframe', background=app_bg, font=bold_font, foreground=text_color)
        self.style.configure('TLabelframe.Label', background=app_bg, font=bold_font, foreground=text_color)
        
        # 啟動動作按鈕：深灰色 (專業感)
        self.style.configure('Action.TButton', font=bold_font, foreground='white', background='#333333', borderwidth=0)
        self.style.map('Action.TButton', background=[('active', '#000000')])
        
        # 一般按鈕：淺灰色
        self.style.configure('Normal.TButton', font=main_font, foreground=text_color, background='#CCCCCC')
        self.style.map('Normal.TButton', background=[('active', '#AAAAAA')])
        
        # 撤銷/儲存/危險按鈕：重點紅色 (保留)
        self.style.configure('Red.TButton', font=bold_font, foreground='white', background='#CC0000', borderwidth=0)
        self.style.map('Red.TButton', background=[('active', '#990000')]) 
        
        self.style.configure('TLabel', background=app_bg, font=main_font, foreground=text_color)
        
        # 標籤頁 (Notebook)：灰白切換
        self.style.configure('TNotebook', background='#D0D0D0')
        self.style.configure('TNotebook.Tab', font=bold_font, padding=[12, 6], background='#D0D0D0', foreground='#444444')
        self.style.map('TNotebook.Tab', background=[('selected', '#E8E8E8')], foreground=[('selected', '#000000')])
        # -------------------

        
        try:
            import sys
            import os
            # PyInstaller 執行時會將檔案解壓縮到 _MEIPASS 暫存資料夾
            base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
            icon_path = os.path.join(base_path, 'icon.ico')
            self.iconbitmap(icon_path)
        except Exception as e:
            pass # 找不到就算了，不影響主程式執行

        self.create_widgets()
        sys.stdout = TextRedirector(self.log_text, "stdout")
        sys.stderr = TextRedirector(self.log_text, "stderr")
        log_msg("=== 歡迎使用 股東e票通 - 輕鬆投自動化助手 ===")
        log_msg("本程式為輔助投票+截圖小幫手，資料皆留存於電腦內，請放心使用")
        log_msg("小提示: 第一次使用請先到 [設定] 頁面輸入帳號喔！")
        
        self.after(200, self.show_disclaimer)
        self.after(100, lambda: check_for_updates(auto=True))
        
    def on_closing(self):
        global main_window_geom
        main_window_geom = self.geometry()
        self.save_config()
        self.destroy()

    def show_disclaimer(self, force_show=False):
        global disclaimer_agreed, disc_window_geom
        if disclaimer_agreed and not force_show: return
            
        disclaimer_win = tk.Toplevel(self)
        disclaimer_win.title("免責聲明與使用提示")
        
        if disc_window_geom:
            disclaimer_win.geometry(disc_window_geom)
        else:
            w, h = 550, 510 
            x = self.winfo_x() + (self.winfo_width() // 2) - (w // 2)
            y = self.winfo_y() + (self.winfo_height() // 2) - (h // 2)
            disclaimer_win.geometry(f"{w}x{h}+{x}+{y}")
            
        disclaimer_win.resizable(True, True)
        disclaimer_win.minsize(400, 400)
        disclaimer_win.transient(self)
        
        def on_disclaimer_close():
            global disc_window_geom
            disc_window_geom = disclaimer_win.geometry()
            self.save_config()
            disclaimer_win.destroy()
        
        if not disclaimer_agreed:
            disclaimer_win.grab_set()
            # 沒同意直接按右上角關閉的話，連主程式一起關
            def on_force_close():
                global disc_window_geom
                disc_window_geom = disclaimer_win.geometry()
                self.save_config()
                self.destroy()
            disclaimer_win.protocol("WM_DELETE_WINDOW", on_force_close)
        else:
            disclaimer_win.protocol("WM_DELETE_WINDOW", on_disclaimer_close)

        container = ttk.Frame(disclaimer_win, padding="20")
        container.pack(fill="both", expand=True)

        title_lbl = ttk.Label(container, text="【 免責聲明與使用提示 】", font=('Microsoft JhengHei', 14, 'bold'), foreground='#0056b3')
        title_lbl.pack(pady=(0, 10))

        def add_section(parent, title, content):
            section_frame = ttk.Frame(parent)
            section_frame.pack(fill="x", pady=4)
            t = tk.Label(section_frame, text=title, font=('Microsoft JhengHei', 11, 'bold'), anchor="w", justify="left")
            t.pack(fill="x")
            c = tk.Label(section_frame, text=content, font=('Microsoft JhengHei', 10), anchor="w", justify="left", wraplength=480)
            c.pack(fill="x", padx=(10, 0))

        add_section(container, "一、 安全保證", "此軟體程式碼為 Python 腳本指令\n所有操作皆在您的「本地電腦端」執行，絕無上傳個人資料或傳輸密碼至網路之行為，請安心使用。")
        add_section(container, "二、 本地儲存", "程式記憶之檔案皆經加密處理並儲存於您的電腦路徑中，詳細路徑可於「設定」頁面查看。")
        
        legal_text = (
            "1. 使用者啟動程式即代表已獲得名單內股東之合法授權。\n"
            "2. 程式將自動點選系統告知條款，使用者視同已閱讀並同意。\n"
            "3. 任何因違反證交法之行為，概由使用者本人承擔全部法律責任。"
        )
        add_section(container, "三、 法律聲明與授權確認", legal_text)

        add_section(container, "四、 首次使用須知", "第一次使用請務必先前往「設定」分頁，輸入身分證字號等資訊並按下儲存，否則自動化任務將無法正常執行。")

        footer_msg = "\n點擊下方「我已同意」按鈕，即代表您已了解並同意上述說明。"
        footer_lbl = ttk.Label(container, text=footer_msg, font=('Microsoft JhengHei', 10, 'italic'), foreground='#666')
        footer_lbl.pack(pady=(0, 0))

        btn_frame = ttk.Frame(disclaimer_win)
        btn_frame.pack(fill="x", pady=(0, 5))
        
        if not disclaimer_agreed:
            def on_agree():
                global disclaimer_agreed
                disclaimer_agreed = True
                on_disclaimer_close()
            agree_btn = ttk.Button(btn_frame, text="我已同意並開始使用", command=on_agree, style='Action.TButton')
            agree_btn.pack(pady=(0, 5), ipadx=20, ipady=5)
        else:
            close_btn = ttk.Button(btn_frame, text="關閉視窗", command=on_disclaimer_close)
            close_btn.pack(pady=(0, 5), ipadx=20)

    def load_config(self):
        global shareholderIDs, vote_speed, shot_speed, screenshot_mode, manual_vote, default_vote, accept_list, opposite_list, abstain_list, login_type, disclaimer_agreed, main_window_geom, disc_window_geom, join_draw, ignore_update_until, saved_sites, browser_choice, last_selected_site
        conf_path = os.path.join(CONFIG_DIR, 'program_setting.conf')
        vote_conf_path = os.path.join(CONFIG_DIR, 'vote_setting.conf')
        
        if not isinstance(saved_sites, dict): saved_sites = {}

        if os.path.exists(conf_path):
            try:
                with open(conf_path, 'r', encoding='utf8') as f:
                    content = f.read()
                    for line in content.split('\n'):
                        if "screenshot_mode:::" in line: screenshot_mode = int(line.split(":::")[1])
                        if "vote_speed:::" in line: vote_speed = float(line.split(":::")[1])
                        if "shot_speed:::" in line: shot_speed = float(line.split(":::")[1])
                        if "login_type:::" in line: login_type = line.split(":::")[1]
                        if "disclaimer_agreed:::" in line: disclaimer_agreed = (line.split(":::")[1].strip() == 'True') 
                        if "main_window_geom:::" in line: main_window_geom = line.split(":::")[1].strip()
                        if "disc_window_geom:::" in line: disc_window_geom = line.split(":::")[1].strip()
                        if "join_draw:::" in line: join_draw = (line.split(":::")[1].strip() == 'True')
                        if "ignore_update_until:::" in line: ignore_update_until = float(line.split(":::")[1])
                        if "browser_choice:::" in line: browser_choice = line.split(":::")[1].strip()
                        if "last_selected_site:::" in line: last_selected_site = line.split(":::")[1].strip() # <--- 讀取上次記憶網頁
                        if "saved_sites:::" in line: 
                            try: saved_sites = json.loads(line.split(":::")[1])
                            except: saved_sites = {}
                        if "shareholderIDs:::" in line: 
                            encrypted_ids = line.split(":::")[1]
                            decrypted_ids = decrypted_ids = decrypt_data(encrypted_ids)
                            if decrypted_ids: shareholderIDs = decrypted_ids.split("|/|")
            except: pass
            
        if os.path.exists(vote_conf_path):
            try:
                with open(vote_conf_path, 'r', encoding='utf8') as f:
                    content = f.read()
                    for line in content.split('\n'):
                        if "default:::" in line: default_vote = line.split(":::")[1]
                        if "manual_vote:::" in line: manual_vote = (line.split(":::")[1] == 'True')
                        if "accept:::" in line: accept_list = [k for k in line.split(":::")[1].split("|/|") if k]
                        if "opposite:::" in line: opposite_list = [k for k in line.split(":::")[1].split("|/|") if k]
                        if "abstain:::" in line: abstain_list = [k for k in line.split(":::")[1].split("|/|") if k]
            except: pass

        self.vote_speed_var = tk.StringVar(value=str(vote_speed))
        self.shot_speed_var = tk.StringVar(value=str(shot_speed))
        self.screenshot_mode_var = tk.IntVar(value=screenshot_mode)
        self.join_draw_var = tk.BooleanVar(value=join_draw) 
        ids_str = ",".join(shareholderIDs) if shareholderIDs else ""
        self.ids_var = tk.StringVar(value=ids_str)
        self.revoke_mode_var = tk.StringVar(value="specific")
        self.browser_choice_var = tk.StringVar(value=browser_choice)

    def save_config(self):
        global vote_speed, shot_speed, screenshot_mode, shareholderIDs, login_type, disclaimer_agreed, main_window_geom, disc_window_geom, join_draw, ignore_update_until, saved_sites, browser_choice, last_selected_site
        try:
            try: 
                v_val = float(self.vote_speed_var.get()); vote_speed = v_val 
                s_val = float(self.shot_speed_var.get()); shot_speed = s_val 
            except ValueError:
                messagebox.showerror("設定錯誤", "速度倍率請輸入有效的數字")
                return
            screenshot_mode = self.screenshot_mode_var.get()
            login_type = self.login_type_cb.get()
            join_draw = self.join_draw_var.get() 
            browser_choice = self.browser_choice_var.get() 
            
            # 取得最後選擇的網頁
            if hasattr(self, 'selected_site'):
                last_selected_site = self.selected_site.get()
                
            shareholderIDs = [x.strip() for x in self.ids_var.get().split(',') if x.strip()]
            conf_path = os.path.join(CONFIG_DIR, 'program_setting.conf')
            with open(conf_path, 'w', encoding='utf8') as f:
                f.write(f"screenshot_mode:::{screenshot_mode}\n")
                f.write(f"vote_speed:::{vote_speed}\n")
                f.write(f"shot_speed:::{shot_speed}\n")
                f.write(f"login_type:::{login_type}\n")
                f.write(f"disclaimer_agreed:::{disclaimer_agreed}\n") 
                f.write(f"main_window_geom:::{main_window_geom}\n")
                f.write(f"disc_window_geom:::{disc_window_geom}\n")
                f.write(f"join_draw:::{join_draw}\n")
                f.write(f"ignore_update_until:::{ignore_update_until}\n")
                f.write(f"browser_choice:::{browser_choice}\n") 
                f.write(f"last_selected_site:::{last_selected_site}\n") # <--- 寫入記憶
                f.write(f"saved_sites:::{json.dumps(saved_sites, ensure_ascii=False)}\n")
                encrypted_str = encrypt_data('|/|'.join(shareholderIDs))
                f.write(f"shareholderIDs:::{encrypted_str}\n")
                f.write("hash:::SECURE_ENCRYPTED_V4\n")
            log_msg("設定已儲存。")
        except Exception as e: log_msg(f"儲存設定失敗: {e}")
    

    # 速率微調輔助函數
    def _adj_val(self, var, delta):
        try:
            # 取得當前數值，若為空或非數字則預設為 1.0
            curr_val = float(var.get()) if var.get() else 1.0
            # 四捨五入到小數點後第一位，避免浮點數運算誤差 (如 1.2000000000000002)
            new_val = round(curr_val + delta, 1)
            # 設定最低門檻 0.1，防止數值歸零或變成負數
            if new_val < 0.1: new_val = 0.1
            var.set(str(new_val))
        except ValueError:
            # 如果使用者手動輸入了非數字，點擊按鈕時不作動或可重設為 1.0
            pass

    def create_widgets(self):
        main_container = ttk.Frame(self, padding="15")
        main_container.pack(fill="both", expand=True)
        
        tab_control = ttk.Notebook(main_container)
        
        tab1 = ttk.Frame(tab_control, padding="10")
        tab2 = ttk.Frame(tab_control, padding="10") 
        tab3 = ttk.Frame(tab_control, padding="10")
        tab4 = ttk.Frame(tab_control, padding="10") 
        tab5 = ttk.Frame(tab_control, padding="10") # <--- 新增 tab5 (網址管理)
        
        tab_control.add(tab1, text='  自動任務  ')
        tab_control.add(tab2, text='  撤銷投票  ')
        tab_control.add(tab3, text='  設定  ')
        tab_control.add(tab4, text='  系統資訊  ') 
        tab_control.add(tab5, text='  網址管理  ') # <--- 加入選單

        # 執行狀態框
        self.frame_log = ttk.LabelFrame(main_container, text=" 執行狀態 ")
        self.log_text = scrolledtext.ScrolledText(self.frame_log, height=8, state='disabled', font=('Consolas', 9), bg='#ffffff', fg='#333333')
        self.log_text.pack(expand=True, fill="both", padx=5, pady=5)

        tab_control.pack(side="top", expand=True, fill="both")
        tab_control.bind("<<NotebookTabChanged>>", self.on_tab_change)

        # ==========================================
        # === Tab 1: 自動任務 ===
        # ==========================================
        frame_mode = ttk.LabelFrame(tab1, text=" 自動化小幫手 ")
        frame_mode.pack(fill="x", pady=(5, 5), ipady=5)
        desc_lbl = ttk.Label(frame_mode, text="說明: 設定好帳號，程式會自動幫您完成投票與截圖存檔，輕鬆領紀念品！")
        desc_lbl.pack(padx=10, pady=(5,5), anchor="w")
        btn_mode1 = ttk.Button(frame_mode, text="啟動程式", style='Action.TButton', command=self.start_mode_1, cursor="hand2")
        btn_mode1.pack(fill="x", padx=15, pady=5)
        
        frame_mode2 = ttk.LabelFrame(tab1, text=" 單筆補圖工具 ")
        frame_mode2.pack(fill="x", pady=5, ipady=5)
        grid_frame = ttk.Frame(frame_mode2)
        grid_frame.pack(fill="x", padx=10, pady=5)
        ttk.Label(grid_frame, text="身分證字號:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.single_id_entry = ttk.Entry(grid_frame, width=40)
        self.single_id_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        
        ttk.Label(grid_frame, text="股票代號:\n(可貼上Excel欄位\n或用逗號區隔)").grid(row=1, column=0, padx=5, pady=5, sticky="nw")
        text_frame = ttk.Frame(grid_frame)
        text_frame.grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        self.stock_list_entry = tk.Text(text_frame, width=40, height=3, font=('Microsoft JhengHei', 10))
        self.stock_list_entry.pack(side="left", fill="both", expand=True)
        scrollbar = ttk.Scrollbar(text_frame, orient="vertical", command=self.stock_list_entry.yview)
        scrollbar.pack(side="right", fill="y")
        self.stock_list_entry.configure(yscrollcommand=scrollbar.set)
        
        grid_frame.columnconfigure(1, weight=1)
        btn_mode2 = ttk.Button(frame_mode2, text="執行單筆補圖", style='Normal.TButton', command=self.start_mode_2, cursor="hand2")
        btn_mode2.pack(fill="x", padx=15, pady=5)

        ttk.Label(frame_mode2, text="💡 小撇步: 身分證欄位留空，系統會自動抓取「設定」裡的所有帳號喔！", foreground="#FF4500", font=('Microsoft JhengHei', 9, 'bold')).pack(pady=(0, 5))

        # ==========================================
        # === Tab 1 新增: 後續動作區 (黑白灰風格) ===
        # ==========================================
        frame_post_action = ttk.LabelFrame(tab1, text=" 任務完成後動作 ")
        frame_post_action.pack(fill="x", pady=(0, 5), ipady=5, padx=2)

        action_inner = ttk.Frame(frame_post_action)
        action_inner.pack(fill="x", padx=10, pady=5)
        
        ttk.Label(action_inner, text="選擇使用的瀏覽器:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        browser_frame = ttk.Frame(action_inner)
        browser_frame.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        ttk.Radiobutton(browser_frame, text="Edge", variable=self.browser_choice_var, value="Edge").pack(side="left", padx=(0, 15))
        ttk.Radiobutton(browser_frame, text="Chrome", variable=self.browser_choice_var, value="Chrome").pack(side="left")

        # --- 換成黑白灰風格的背景 ---
        pop_bg = '#DCDCDC' # 灰階底色
        combo_row_frame = tk.Frame(action_inner, bg=pop_bg, padx=5, pady=4)
        combo_row_frame.grid(row=1, column=0, columnspan=2, sticky="ew", pady=5)

        tk.Label(combo_row_frame, text="✨ 任務完畢後開啟:", bg=pop_bg, font=('Microsoft JhengHei', 10, 'bold'), fg='#000000').pack(side="left", padx=5)
        
        # 綁定全域記憶變數 last_selected_site
        self.selected_site = tk.StringVar(value=last_selected_site)
        self.site_combo = ttk.Combobox(combo_row_frame, textvariable=self.selected_site, state="readonly", width=35)
        self.site_combo.pack(side="left", padx=5, fill="x", expand=True)
        # --- 區塊結束 ---
        
        action_inner.columnconfigure(1, weight=1)
        
        # ==========================================
        # === Tab 2: 撤銷投票 ===
        # ==========================================
        frame_revoke = ttk.LabelFrame(tab2, text=" 撤銷投票小幫手 ")
        frame_revoke.pack(fill="x", pady=5, ipady=5)
        
        r_grid = ttk.Frame(frame_revoke)
        r_grid.pack(fill="x", padx=10, pady=5)
        
        ttk.Label(r_grid, text="身分證字號:").grid(row=0, column=0, padx=5, pady=7, sticky="w")
        self.revoke_id_entry = ttk.Entry(r_grid, width=40)
        self.revoke_id_entry.grid(row=0, column=1, padx=5, pady=7, sticky="ew")
        ttk.Label(r_grid, text="(留空則載入設定帳號)", foreground="#888").grid(row=0, column=2, padx=5, sticky="w")
        
        ttk.Label(r_grid, text="撤銷模式:").grid(row=1, column=0, padx=5, pady=7, sticky="nw")
        mode_frame = ttk.Frame(r_grid)
        mode_frame.grid(row=1, column=1, padx=5, pady=5, sticky="w", columnspan=2)
        ttk.Radiobutton(mode_frame, text="指定代號撤銷", variable=self.revoke_mode_var, value="specific").pack(side="left", padx=(0, 15))
        ttk.Radiobutton(mode_frame, text="全部撤銷 (掃描列表)", variable=self.revoke_mode_var, value="all").pack(side="left")
        
        ttk.Label(r_grid, text="指定股票代號:\n(全部撤銷免填)").grid(row=2, column=0, padx=5, pady=8, sticky="nw")
        
        r_text_frame = ttk.Frame(r_grid)
        r_text_frame.grid(row=2, column=1, padx=6, pady=5, sticky="ew")
        
        self.revoke_stock_list = tk.Text(r_text_frame, width=33, height=8, font=('Microsoft JhengHei', 10))
        self.revoke_stock_list.pack(side="left", fill="both", expand=True)
        
        r_scroll = ttk.Scrollbar(r_text_frame, orient="vertical", command=self.revoke_stock_list.yview)
        r_scroll.pack(side="right", fill="y")
        
        self.revoke_stock_list.configure(yscrollcommand=r_scroll.set)
        
        r_grid.columnconfigure(1, weight=1)
        btn_revoke = ttk.Button(frame_revoke, text="開始撤銷任務", style='Red.TButton', command=self.start_mode_3, cursor="hand2")
        btn_revoke.pack(fill="x", padx=15, pady=10)
        
        ttk.Label(frame_revoke, text="⚠️ 提示：若電腦有多人憑證，跳出視窗時請務必選擇「對應目前帳號身分證」的憑證！\n完成後程式會自動偵測並繼續下一筆。", foreground="#d9534f", justify="center", font=('Microsoft JhengHei', 9, 'bold')).pack(pady=(0, 10))

        # ==========================================
        # === Tab 3: 設定 ===
        # ==========================================
        frame_setting = ttk.Frame(tab3)
        frame_setting.pack(fill="both", expand=True, padx=10, pady=12)
        
        type_frame = ttk.Frame(frame_setting)
        type_frame.pack(fill="x", pady=2)
        ttk.Label(type_frame, text="請選擇登入驗證方式:").pack(anchor="w")
        login_types = ["券商網路下單憑證", "自然人憑證", "行動自然人憑證"]
        self.login_type_cb = ttk.Combobox(type_frame, values=login_types, state="readonly", width=25)
        self.login_type_cb.set(login_type if login_type in login_types else "券商網路下單憑證")
        self.login_type_cb.pack(anchor="w", pady=8)
        
        ttk.Separator(frame_setting, orient='horizontal').pack(fill='x', pady=12)

        spd_frame = ttk.Frame(frame_setting)
        spd_frame.pack(fill="x", pady=2)
        
        v_frame = ttk.Frame(spd_frame)
        v_frame.pack(fill="x", pady=2)
        ttk.Label(v_frame, text="投票速度：", width=20).pack(side="left")
        ttk.Button(v_frame, text="-", width=1, command=lambda: self._adj_val(self.vote_speed_var, -0.1)).pack(side="left", padx=2)
        ttk.Entry(v_frame, textvariable=self.vote_speed_var, width=8, justify='center').pack(side="left", padx=2)
        ttk.Button(v_frame, text="+", width=1, command=lambda: self._adj_val(self.vote_speed_var, 0.1)).pack(side="left", padx=2)
        ttk.Label(v_frame, text="(0.1=極速, 10=慢速)", foreground="#666").pack(side="left", padx=5)

        s_frame = ttk.Frame(spd_frame)
        s_frame.pack(fill="x", pady=2)
        ttk.Label(s_frame, text="截圖速度：", width=20).pack(side="left")
        ttk.Button(s_frame, text="-", width=1, command=lambda: self._adj_val(self.shot_speed_var, -0.1)).pack(side="left", padx=2)
        ttk.Entry(s_frame, textvariable=self.shot_speed_var, width=8, justify='center').pack(side="left", padx=2)
        ttk.Button(s_frame, text="+", width=1, command=lambda: self._adj_val(self.shot_speed_var, 0.1)).pack(side="left", padx=2)
        ttk.Label(s_frame, text="(0.1=極速, 10=慢速)", foreground="#666").pack(side="left", padx=5)
        
        ttk.Label(spd_frame, text="💡 提醒：若無法輸入小數點，請切換至英文輸入法", foreground="#d9534f").pack(anchor="w", pady=(5, 0))
        ttk.Separator(frame_setting, orient='horizontal').pack(fill='x', pady=12)
        
        file_frame = ttk.Frame(frame_setting)
        file_frame.pack(fill="x", pady=2)
        ttk.Label(file_frame, text="截圖存檔方式:").pack(anchor="w", pady=(0,2))
        ttk.Radiobutton(file_frame, text="A. 每個帳號獨立資料夾", variable=self.screenshot_mode_var, value=1).pack(anchor="w", padx=10)
        ttk.Radiobutton(file_frame, text="B. 全部放在一起同一資料夾(檔名會加上戶名)", variable=self.screenshot_mode_var, value=2).pack(anchor="w", padx=10)
        ttk.Separator(frame_setting, orient='horizontal').pack(fill='x', pady=12)

        draw_frame = ttk.Frame(frame_setting)
        draw_frame.pack(fill="x", pady=2)
        tk.Checkbutton(draw_frame, text="遇到抽獎頁面時，暫停 5 分鐘讓我手動參加抽獎", variable=self.join_draw_var, bg='#F5F7FA', activebackground='#F5F7FA', font=('Microsoft JhengHei', 10)).pack(anchor="w")
        ttk.Label(draw_frame, text="(若未勾選則程式會自動關閉視窗略過)", foreground="#666").pack(anchor="w", padx=20)
        ttk.Separator(frame_setting, orient='horizontal').pack(fill='x', pady=12)
        
        id_frame = ttk.Frame(frame_setting)
        id_frame.pack(fill="x", pady=2)
        ttk.Label(id_frame, text="我的帳號清單 (多個請用逗號分隔):").pack(anchor="w")
        ttk.Entry(id_frame, textvariable=self.ids_var).pack(fill="x", pady=1)
        ttk.Button(frame_setting, text="儲存設定 (加密)", style='Red.TButton', command=self.save_config, cursor="hand2").pack(pady=2, ipady=1, fill='x')
        ttk.Button(frame_setting, text="檢查程式更新", style='Action.TButton', command=lambda: check_for_updates(auto=False)).pack(pady=5, ipady=1, fill='x')

        # ==========================================
        # === Tab 4: 系統資訊 ===
        # ==========================================
        frame_info = ttk.Frame(tab4)
        frame_info.pack(fill="both", expand=True, padx=10, pady=12)

        path_info_frame = ttk.LabelFrame(frame_info, text=" 📂 資料夾路徑 (可全選複製) ")
        path_info_frame.pack(fill="x", pady=(0, 15))

        ttk.Label(path_info_frame, text="設定檔記憶位置:", foreground="#555").pack(anchor="w", padx=10, pady=(5,0))
        path_entry_conf = ttk.Entry(path_info_frame)
        path_entry_conf.insert(0, CONFIG_DIR) 
        path_entry_conf.configure(state="readonly") 
        path_entry_conf.pack(fill="x", padx=10, pady=5)

        ttk.Label(path_info_frame, text="截圖存檔位置:", foreground="#555").pack(anchor="w", padx=10, pady=(5,0))
        path_entry_shot = ttk.Entry(path_info_frame)
        abs_shot_path = os.path.abspath(base_path) 
        path_entry_shot.insert(0, abs_shot_path)
        path_entry_shot.configure(state="readonly")
        path_entry_shot.pack(fill="x", padx=10, pady=(0,10))

        disc_frame = ttk.LabelFrame(frame_info, text=" 📜 授權與聲明 ")
        disc_frame.pack(fill="x", pady=(0, 5))
        ttk.Label(disc_frame, text="本程式僅供個人輔助使用，無任何資料上傳行為。\n第一次啟動時已確認授權，您也可隨時重新閱讀。", justify="center").pack(pady=10)
        btn_disclaimer = ttk.Button(disc_frame, text="📝 重新閱讀免責聲明", command=lambda: self.show_disclaimer(force_show=True))
        btn_disclaimer.pack(pady=(0, 10), ipadx=10)

        # ==========================================
        # === Tab 5: 網址管理 (獨立設定與刪除) ===
        # ==========================================
        frame_url = ttk.Frame(tab5)
        frame_url.pack(fill="both", expand=True, padx=10, pady=12)

        manage_frame = ttk.LabelFrame(frame_url, text=" 📝 新增完成後開啟的網站 ")
        manage_frame.pack(fill="x", pady=5, ipady=10)

        grid_url = ttk.Frame(manage_frame)
        grid_url.pack(fill="x", padx=10, pady=5)
        
        ttk.Label(grid_url, text="自訂名稱:").grid(row=0, column=0, padx=10, pady=5, sticky="w")
        self.entry_name = ttk.Entry(grid_url)
        self.entry_name.grid(row=0, column=1, padx=10, pady=5, sticky="ew")

        ttk.Label(grid_url, text="網址 (URL):").grid(row=1, column=0, padx=10, pady=5, sticky="w")
        self.entry_url = ttk.Entry(grid_url)
        self.entry_url.grid(row=1, column=1, padx=10, pady=5, sticky="ew")

        btn_frame = ttk.Frame(manage_frame)
        btn_frame.pack(pady=10)
        ttk.Button(btn_frame, text="儲存 / 新增", command=self.add_site).pack(side="left", padx=10)
        grid_url.columnconfigure(1, weight=1)

        del_frame = ttk.LabelFrame(frame_url, text=" 🗑️ 刪除已儲存的網站 ")
        del_frame.pack(fill="x", pady=15, ipady=10)
        
        del_inner = ttk.Frame(del_frame)
        del_inner.pack(fill="x", padx=10, pady=10)
        ttk.Label(del_inner, text="選擇要刪除的網站:").pack(side="left", padx=(10, 5))
        
        self.manage_selected_site = tk.StringVar()
        self.manage_site_combo = ttk.Combobox(del_inner, textvariable=self.manage_selected_site, state="readonly", width=30)
        self.manage_site_combo.pack(side="left", padx=10)
        ttk.Button(del_inner, text="刪除此網站", command=self.delete_site, style='Red.TButton').pack(side="left", padx=10)

        ttk.Label(frame_url, text="💡 設定好的網址可以在「自動任務」分頁底下的藍色區塊挑選！", foreground="#666").pack(pady=10)

        # === 初始化排版魔法 ===
        self.frame_log.pack(in_=tab1, fill="both", expand=True, pady=(5, 0))
        self.update_site_list() # 初始化網址清單

    # ================= 新增：網址管理與後續動作 =================
    def update_site_list(self):
        site_names = list(saved_sites.keys())
        
        # 1. 更新首頁下拉選單 (加入不開啟選項)
        tab1_options = ["【不開啟任何網頁】"] + site_names
        self.site_combo['values'] = tab1_options
        
        # 若上次儲存的選項在列表中，則選中它，否則預設不開啟
        if last_selected_site in tab1_options:
            self.selected_site.set(last_selected_site)
        elif self.selected_site.get() not in tab1_options:
            self.selected_site.set("【不開啟任何網頁】")
            
        # 2. 更新管理頁面下拉選單 (純資料清單)
        self.manage_site_combo['values'] = site_names
        if site_names:
            if self.manage_selected_site.get() not in site_names:
                self.manage_selected_site.set(site_names[0])
        else:
            self.manage_selected_site.set("")

    def add_site(self):
        name = self.entry_name.get().strip()
        url = self.entry_url.get().strip()
        if not name or not url:
            messagebox.showwarning("提示", "名稱和網址都不能是空的喔！")
            return
        saved_sites[name] = url
        self.save_config()
        self.update_site_list()
        
        # 新增完自動切換
        self.selected_site.set(name) 
        self.manage_selected_site.set(name)
        
        self.entry_name.delete(0, tk.END)
        self.entry_url.delete(0, tk.END)
        log_msg(f"✅ 已儲存網站: {name}")

    def delete_site(self):
        target = self.manage_selected_site.get()
        if not target: 
            messagebox.showwarning("提示", "請先選擇要刪除的網站！")
            return
        if messagebox.askyesno("確認", f"確定要在資料庫中刪除「{target}」嗎？"):
            if target in saved_sites:
                del saved_sites[target]
                self.save_config()
                self.update_site_list()
                log_msg(f"🗑️ 已刪除網站: {target}")

    # ================= 新增：網址管理與智慧彈跳控制 =================
    def _finish_task(self):
        # 判斷本次是否有成功截圖的項目
        total_screenshots = sum(len(res.get('success_screenshot', [])) for res in session_results.values())
        has_shots = total_screenshots > 0

        if has_shots:
            # 1. 跑完報告後先開啟網頁
            target_name = self.selected_site.get()
            browser = self.browser_choice_var.get()
            
            if target_name != "【不開啟任何網頁】" and target_name:
                url = saved_sites.get(target_name, "")
                if url:
                    log_msg(f"準備使用 {browser} 開啟設定網頁: {target_name}")
                    try:
                        if browser == "Edge": subprocess.Popen(f'start msedge "{url}"', shell=True)
                        elif browser == "Chrome": subprocess.Popen(f'start chrome "{url}"', shell=True)
                    except Exception as e:
                        log_msg(f"開啟瀏覽器時發生錯誤: {e}")

            # 2. 開完網頁後，間隔 1 秒 (1000毫秒) 再開啟資料夾
            self.after(1000, self._open_folder_and_notify)
        else:
            # 如果沒做任何截圖動作，就不用開視窗了
            log_msg("本次無執行截圖，不開啟設定網頁與資料夾。")
            self._pop_topmost_message("任務搞定！報告已經產生！\n\n(提示：本次無執行截圖，故未開啟網頁與資料夾)")

    def _open_folder_and_notify(self):
        abs_path = os.path.abspath(base_path)
        log_msg("正在為您開啟截圖資料夾...")
        try:
            # 這時候開資料夾，會自動疊在剛剛開好的網頁上方
            subprocess.Popen(f'explorer "{abs_path}"')
        except Exception as e:
            log_msg(f"開啟截圖資料夾失敗: {e}")

        # 等資料夾視窗出現 (給系統1.2秒反應時間)，再進行排版
        self.after(1200, self._resize_folder_and_notify, abs_path)

    def _resize_folder_and_notify(self, abs_path):
        import ctypes
        try:
            folder_name = os.path.basename(abs_path)
            # 尋找 Windows 檔案總管視窗
            hwnd = ctypes.windll.user32.FindWindowW("CabinetWClass", folder_name)
            if not hwnd:
                hwnd = ctypes.windll.user32.FindWindowW(None, folder_name)

            if hwnd:
                # 取得螢幕解析度
                sw = ctypes.windll.user32.GetSystemMetrics(0)
                sh = ctypes.windll.user32.GetSystemMetrics(1)
                
                # 設定為 1/4 大小 (寬一半、高一半)
                w, h = sw // 2, sh // 2
                x, y = 50, 50 
                
                # 關鍵修正：-1 代表 HWND_TOPMOST (強制無敵置頂，連 Chrome 都蓋不掉)
                # 0x0040 = SWP_SHOWWINDOW
                ctypes.windll.user32.SetWindowPos(hwnd, -1, x, y, w, h, 0x0040)
                ctypes.windll.user32.SetForegroundWindow(hwnd)
                
                # 給予 0.5 秒的強制置頂時間對抗 Chrome，然後解除強制置頂 (-2 代表 HWND_NOTOPMOST)
                # 這樣資料夾能停留在網頁上方，且您後續點擊網頁時，網頁還是可以正常蓋過資料夾
                self.after(500, lambda: ctypes.windll.user32.SetWindowPos(hwnd, -2, x, y, w, h, 0x0040))
                
        except Exception as e:
            log_msg(f"資料夾視窗自動排版失敗: {e}")

        # 資料夾排版完成後，再間隔 1 秒 (1000毫秒)，最後才跳出程式完成提示
        self.after(1000, self._pop_topmost_message, "任務搞定！報告已經產生！\n\n網頁與資料夾已為您開啟。")

    def _pop_topmost_message(self, msg):
        # 將主程式拉到最上層
        self.attributes('-topmost', True)
        self.update()
        # 關鍵：馬上解除最上層鎖定，讓彈跳視窗不會死鎖在螢幕前
        self.attributes('-topmost', False)
        
        # 統計失敗數量 (包含投票失敗與截圖失敗)
        fail_count = sum(len(res.get('fail_vote', [])) + len(res.get('fail_screenshot', [])) for res in session_results.values())

        if fail_count > 0:
            # 如果有失敗，將原本預設的成功字眼換掉，並加上警告提示
            msg = msg.replace("任務搞定！", "任務執行結束！")
            warning_msg = f"{msg}\n\n⚠️ 注意：偵測到 {fail_count} 筆「投票」或「截圖」失敗！\n👉 請務必查看 LOG 報告檔確認詳細失敗原因，並視情況重新跑一次。"
            messagebox.showwarning("任務結束 (有部分失敗)", warning_msg)
        else:
            # 完全成功的情況，維持原本的提示
            messagebox.showinfo("完成", msg)

    def on_tab_change(self, event):
        selected_tab = event.widget.select()
        tab_text = event.widget.tab(selected_tab, "text").strip()

        # 切換分頁時，先解開目前的排版
        self.frame_log.pack_forget()

        # 如果點擊的是這兩頁，就動態把 Log 框放進去，並設定 expand=True 把底下的空白填滿！
        if "自動任務" in tab_text or "撤銷投票" in tab_text:
            target_tab = event.widget.nametowidget(selected_tab)
            self.frame_log.pack(in_=target_tab, fill="both", expand=True, pady=(5, 0))
            
    def start_mode_1(self):
        self.save_config()
        self.log_text.configure(state='normal')
        self.log_text.delete(1.0, tk.END)
        self.log_text.configure(state='disabled')
        threading.Thread(target=self.run_logic_mode_1, daemon=True).start()

    def start_mode_2(self):
        input_ids = self.single_id_entry.get().strip()
        stocks = self.stock_list_entry.get("1.0", tk.END).strip()
        target_id_list = []
        if not input_ids:
            self.save_config() 
            target_id_list = [x for x in shareholderIDs if x]
            if not target_id_list:
                messagebox.showerror("錯誤", "設定裡沒有帳號，請手動輸入或去設定新增")
                return
        else: target_id_list = [x.strip() for x in input_ids.split(',') if x.strip()]
        if not stocks:
            messagebox.showerror("輸入錯誤", "請輸入股票代號")
            return
        self.log_text.configure(state='normal')
        self.log_text.delete(1.0, tk.END)
        self.log_text.configure(state='disabled')
        threading.Thread(target=self.run_logic_mode_2, args=(target_id_list, stocks), daemon=True).start()

    def start_mode_3(self):
        input_ids = self.revoke_id_entry.get().strip()
        mode = self.revoke_mode_var.get()
        stocks_str = self.revoke_stock_list.get("1.0", tk.END).strip()
        stock_list = re.findall(r'\d+', stocks_str)
        
        if mode == "all" and stock_list:
            messagebox.showerror("輸入衝突", "您選擇了「全部撤銷」，請將下方的「指定股票代號」欄位清空，或改選「指定代號撤銷」。")
            return
        
        target_id_list = []
        if not input_ids:
            self.save_config() 
            target_id_list = [x for x in shareholderIDs if x]
            if not target_id_list:
                messagebox.showerror("錯誤", "設定裡沒有帳號，請手動輸入或去設定新增")
                return
        else: target_id_list = [x.strip() for x in input_ids.split(',') if x.strip()]
        
        if mode == "specific" and not stock_list:
            messagebox.showerror("輸入錯誤", "選擇「指定代號撤銷」時，請輸入股票代號")
            return
            
        self.log_text.configure(state='normal')
        self.log_text.delete(1.0, tk.END)
        self.log_text.configure(state='disabled')
        threading.Thread(target=self.run_logic_mode_3, args=(target_id_list, mode, stock_list), daemon=True).start()

    def run_logic_mode_1(self):
        global driver, voteinfolist, shareholderIDs, session_results, user_name_map
        session_results = {}; user_name_map = {} 
        start_time = time.time(); total_items = 0
        log_msg("=== 自動任務開始 ===")
        if not os.path.exists(base_path): os.makedirs(base_path)
        
        read_voteinfolist(voteinfolist)
        for uid in shareholderIDs:
             if uid not in voteinfolist: voteinfolist[uid] = {}
             elif isinstance(voteinfolist[uid], list):
                 voteinfolist[uid] = {sid: 0 for sid in voteinfolist[uid]}
                 
        maintenance_flag = False
        
        try:
            if driver is None: driver = get_driver()
        except Exception as e:
            log_msg(f"瀏覽器初始啟動失敗: {e}")
            return

        try:
            for user_id in shareholderIDs:
                if maintenance_flag: break
                if not user_id: continue
                
                try: driver.current_url
                except:
                    log_msg("瀏覽器意外關閉，正在重新啟動...")
                    try: driver = get_driver()
                    except: log_msg("重啟失敗，跳過此帳號"); continue

                try:
                    log_msg(f"--- 正在處理: {user_id} ---")
                    try: autoLogin(user_id)
                    except SystemMaintenanceError:
                        log_msg("!!! 系統維護中，暫停任務 !!!")
                        maintenance_flag = True; break 
                    except LoginTimeoutError: 
                        log_msg("登入逾時或憑證卡死，將重啟瀏覽器...")
                        if user_id not in session_results:
                            session_results[user_id] = {'success': [], 'fail_vote': [], 'fail_screenshot': [], 'success_screenshot': []}
                        session_results[user_id]['fail_vote'].append("登入失敗(逾時/卡死)")
                        force_quit_driver(driver) 
                        driver = None 
                        continue

                    pending_dict = voteinfolist.get(user_id, {})
                    if pending_dict:
                        log_msg(f"發現上次遺留的截圖任務，共 {len(pending_dict)} 張，優先處理重試...")
                        for stock_id in list(pending_dict.keys()):
                            total_items += 1
                            if auto_screenshot(user_id, stock_id) == 0:
                                del voteinfolist[user_id][stock_id] 
                                write_voteinfolist(voteinfolist)
                            else:
                                voteinfolist[user_id][stock_id] += 1 
                                if voteinfolist[user_id][stock_id] > 1:
                                    log_msg(f"[{stock_id}] 截圖已重試失敗，放棄後續重試，避免鬼打牆。")
                                    del voteinfolist[user_id][stock_id]
                                write_voteinfolist(voteinfolist)

                    if autovote(user_id) != 0: log_msg(f"{user_id} 投票過程有點小問題")
                    
                    current_dict = voteinfolist.get(user_id, {})
                    new_items = {k: v for k, v in current_dict.items() if v == 0}
                    
                    if new_items:
                        log_msg(f"開始進行新投票的截圖，共 {len(new_items)} 張...")
                        for stock_id in list(new_items.keys()):
                            total_items += 1
                            if auto_screenshot(user_id, stock_id) == 0:
                                del voteinfolist[user_id][stock_id]
                                write_voteinfolist(voteinfolist)
                            else:
                                voteinfolist[user_id][stock_id] += 1
                                write_voteinfolist(voteinfolist)
                    elif not pending_dict: 
                        log_msg("沒有需要截圖的項目。")
                    
                    logout()
                    
                except Exception as e:
                    log_msg(f"帳號 {user_id} 執行遇到狀況: {e}")
                    force_quit_driver(driver)
                    driver = None

        finally:
            force_quit_driver(driver)
            driver = None

        end_time = time.time()
        generate_session_report(start_time, end_time, total_items)
        log_msg("=== 任務全部完成 ===")
        if maintenance_flag: 
            # 遇到系統維護也彈一下，避免被蓋住
            self.attributes('-topmost', True)
            self.update()
            self.attributes('-topmost', False)
            messagebox.showwarning("暫停", "因為系統維護，目前已終止任務！")
        else: 
            # 觸發全新的 UI 動線
            self.after(0, self._finish_task)

    def run_logic_mode_2(self, id_list, stocks_str):
        global driver, session_results, user_name_map, login_type
        session_results = {}; user_name_map = {} 
        start_time = time.time(); total_items = 0
        
        log_msg(f"=== 開始補圖，共 {len(id_list)} 個帳號 ===")
        stock_list = re.findall(r'\d+', stocks_str)
        if not stock_list:
            log_msg("沒有有效的股票代號")
            return
            
        maintenance_flag = False
        try:
            if driver is None: driver = get_driver()
        except: 
            log_msg("瀏覽器啟動失敗"); return

        try:
            for target_id in id_list:
                if maintenance_flag: break
                try: driver.current_url
                except:
                    log_msg("瀏覽器意外關閉，重啟中...")
                    try: driver = get_driver()
                    except: continue

                try:
                    log_msg(f"--- 補圖帳號: {target_id} (使用: {login_type}) ---")
                    try: 
                        autoLogin(target_id)
                    except SystemMaintenanceError:
                        log_msg("系統維護中，停止任務")
                        maintenance_flag = True; break
                    except LoginTimeoutError: 
                        log_msg("登入逾時，將重啟瀏覽器...")
                        session_results.setdefault(target_id, {'success': [], 'fail_vote': [], 'fail_screenshot': [], 'success_screenshot': []})
                        session_results[target_id]['fail_vote'].append("登入失敗")
                        force_quit_driver(driver)
                        driver = None
                        continue

                    for stock_id in stock_list:
                        total_items += 1
                        auto_screenshot(target_id, stock_id)
                    
                    logout()
                except Exception as e:
                    log_msg(f"{target_id} 執行失敗: {e}")
                    force_quit_driver(driver)
                    driver = None
        finally:
            force_quit_driver(driver)
            driver = None

        end_time = time.time()
        generate_session_report(start_time, end_time, total_items)
        log_msg("=== 補圖任務結束 ===")
        
        # 觸發全新的 UI 動線
        self.after(0, self._finish_task)
        #messagebox.showinfo("完成", "截圖任務完畢")

    def run_logic_mode_3(self, id_list, mode, stock_list):
        global driver, session_results, user_name_map, login_type
        session_results = {}; user_name_map = {} 
        start_time = time.time() # <--- 新增：開始計時
        
        log_msg(f"=== 開始撤銷任務 ({'全部' if mode=='all' else '指定'})，共 {len(id_list)} 個帳號 ===")
        
        maintenance_flag = False
        try:
            if driver is None: driver = get_driver()
        except: 
            log_msg("瀏覽器啟動失敗"); return

        try:
            for target_id in id_list:
                if maintenance_flag: break
                try: driver.current_url
                except:
                    log_msg("瀏覽器意外關閉，重啟中...")
                    try: driver = get_driver()
                    except: continue

                try:
                    log_msg(f"--- 撤銷帳號: {target_id} (使用: {login_type}) ---")
                    try: 
                        autoLogin(target_id)
                    except SystemMaintenanceError:
                        log_msg("系統維護中，停止任務")
                        maintenance_flag = True; break
                    except LoginTimeoutError: 
                        log_msg("登入逾時，將重啟瀏覽器...")
                        force_quit_driver(driver)
                        driver = None
                        continue

                    auto_revoke(target_id, mode, stock_list)
                    logout()
                except Exception as e:
                    log_msg(f"{target_id} 執行失敗: {e}")
                    force_quit_driver(driver)
                    driver = None
        finally:
            force_quit_driver(driver)
            driver = None

        end_time = time.time() # <--- 新增：結束計時
        total_items = sum(len(res.get('success', [])) for res in session_results.values()) # <--- 計算總撤銷數
        
        # --- 新增：產生報告 ---
        generate_session_report(start_time, end_time, total_items)
        
        log_msg("=== 撤銷任務結束 ===")
        
        # 讓撤銷完畢的彈窗也維持在最上層，不會被蓋住
        self.attributes('-topmost', True)
        self.update()
        self.attributes('-topmost', False)
        messagebox.showinfo("完成", "撤銷任務執行完畢，報告已產生！")

if __name__ == "__main__":
    app = App()
    app.mainloop()
