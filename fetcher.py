#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
fetcher.py — 全自動資料抓取
Step1-3: Excel + CMoney 增益集 → 持股明細
Step4:   CMoney 小綠 → 基金規模
Step5:   AI Studio → 分析 + 下載圖表
"""

import os, re, json, time, subprocess, traceback
from datetime import datetime, timedelta
from pathlib import Path

import win32com.client
import win32gui
import win32process
from pywinauto import Application, Desktop
from pywinauto.keyboard import send_keys

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from webdriver_manager.chrome import ChromeDriverManager
import pyperclip

import logger

BASE_DIR = Path(__file__).parent
with open(BASE_DIR / "config.json", encoding="utf-8") as f:
    CFG = json.load(f)

CMONEY_EXE    = r"C:\Program Files (x86)\CMoney\CMoney.exe"
AI_STUDIO_URL = (
    "https://aistudio.google.com/apps/drive/"
    "1DTSmQK6sG0BYB5PfRurcFUWT4R72qd4m?showPreview=true&showAssistant=true"
)
PROFILE_DIR  = BASE_DIR / "browser_profile"
DOWNLOAD_DIR = BASE_DIR / "downloads"
ETF = CFG["etf_code"]


def _p(msg: str, lvl: str = "INFO", cb=None):
    logger.write(msg, lvl)
    if cb:
        try:
            cb(msg, lvl.lower())
        except TypeError:
            cb(msg)


# ══════════════════════════════════════════════════════
#  Step1-3: Excel + CMoney 增益集
# ══════════════════════════════════════════════════════

def fetch_holdings_from_excel(on_progress=None) -> str:
    p = lambda m, lv="INFO": _p(m, lv, on_progress)
    p("啟動 Excel...")

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    wb = excel.Workbooks.Add()
    ws = wb.Sheets(1)
    ws.Name = "今天的00981A"
    ws.Activate()
    time.sleep(1)

    p("設定 CMoney 增益集（00981A操作日報）...")
    _trigger_cmoney_addin(excel, on_progress)

    p("等待資料撈取完成（20秒）...")
    time.sleep(20)

    p("讀取 Excel 資料...")
    rows = _read_sheet_data(excel.ActiveSheet)

    excel.DisplayAlerts = False
    try:
        wb.Close(SaveChanges=False)
    except Exception:
        pass

    p(f"Excel 資料完成：{len(rows)} 筆", "OK")
    return "\n".join(rows)


def _trigger_cmoney_addin(excel, on_progress=None):
    p = lambda m, lv="INFO": _p(m, lv, on_progress)

    hwnd = _hwnd_by_class("XLMAIN")
    if not hwnd:
        raise RuntimeError("找不到 Excel 視窗（XLMAIN）")

    app = Application(backend="uia").connect(handle=hwnd)
    win = app.window(handle=hwnd)
    win.set_focus()
    win.maximize()
    time.sleep(1.5)

    p("嘗試切換至「增益集」標籤...")
    clicked_group = False

    for attempt in range(3):
        try:
            tabs = win.descendants(title="增益集", control_type="TabItem")
            if tabs:
                try:
                    tabs[0].select()
                except Exception:
                    tabs[0].click_input()

            time.sleep(1.5)

            grps = win.descendants(title="自訂工具列", control_type="Group")
            if grps:
                btns = grps[0].descendants(control_type="Button")
                if btns:
                    p("成功深挖到增益集按鈕，點擊...", "INFO")
                    btns[0].click_input()
                    clicked_group = True
                    break
                else:
                    all_items = grps[0].descendants()
                    clickable = [i for i in all_items if i.is_enabled() and i.rectangle().width() > 10]
                    if clickable:
                        clickable[0].click_input()
                        clicked_group = True
                        break
            else:
                p(f"第 {attempt+1} 次尋找自訂工具列失敗，重試中...", "WARN")

        except Exception as e:
            p(f"切換發生異常: {e}", "WARN")
            time.sleep(1)

    if not clicked_group:
        raise RuntimeError("無法成功切換到增益集標籤或找不到按鈕")

    time.sleep(2)
    _handle_cmoney_dialog(on_progress)


def _handle_cmoney_dialog(on_progress=None):
    p = lambda m, lv="INFO": _p(m, lv, on_progress)

    dlg = None
    for _ in range(15):
        try:
            dlgs = Desktop(backend="uia").windows(title_re=r".*資料轉出精靈.*|.*CMoneyExcel.*")
            if dlgs:
                dlg = dlgs[0]
                break
        except Exception:
            pass
        time.sleep(1)

    if not dlg:
        raise RuntimeError("主對話框 (資料轉出精靈) 未出現")

    dlg.set_focus()
    time.sleep(0.5)

    p("尋找「下一步>」按鈕...")
    try:
        btns = dlg.descendants(control_type="Button")
        for b in btns:
            if "下一步" in b.window_text():
                b.click_input()
                p("點擊 下一步>...", "INFO")
                time.sleep(2)
                break
    except Exception as e:
        p(f"點擊下一步時發生錯誤: {e}", "WARN")

    # 檢查子視窗是否已存在
    sub_dlgs = Desktop(backend="uia").windows(title_re=r".*開啟自訂報表.*|.*自訂報表.*")
    if not sub_dlgs:
        p("點擊 開啟...")
        try:
            dlgs = Desktop(backend="uia").windows(title_re=r".*資料轉出精靈.*|.*CMoneyExcel.*")
            if dlgs:
                dlg = dlgs[0]
                dlg.set_focus()
                btns = dlg.descendants(control_type="Button")
                opened = False
                for b in btns:
                    if "開啟" in b.window_text():
                        b.click_input()
                        opened = True
                        break
                if not opened:
                    raise RuntimeError("畫面上找不到「開啟...」按鈕")
        except Exception as e:
            raise RuntimeError(f"尋找「開啟...」按鈕失敗: {e}")
        time.sleep(1.5)

    p("選擇 00981A 操作日報...")
    _select_cmoney_template()
    time.sleep(1)

    p("主對話框 → 確定...")
    try:
        dlgs = Desktop(backend="uia").windows(title_re=r".*資料轉出精靈.*|.*CMoneyExcel.*")
        if dlgs:
            dlg = dlgs[0]
            dlg.set_focus()
            time.sleep(0.3)
            btns = dlg.descendants(control_type="Button")
            for b in btns:
                if b.window_text() == "確定":
                    b.click_input()
                    break
    except Exception as e:
        p(f"點擊主對話框確定失敗: {e}", "WARN")
        send_keys("{ENTER}")

    time.sleep(5)


def _select_cmoney_template():
    # 1. 等待對話框出現
    sub_hwnd = 0
    for _ in range(15):
        sub_hwnd = win32gui.FindWindow(None, "開啟自訂報表")
        if sub_hwnd:
            break
        time.sleep(1)

    if not sub_hwnd:
        raise RuntimeError("等了 15 秒，系統底層依然找不到「開啟自訂報表」對話框")

    # 2. 連接視窗 (加入 try-except 忽略 SetForegroundWindow 報錯)
    app32 = Application(backend="win32").connect(handle=sub_hwnd)
    win32_dlg = app32.window(handle=sub_hwnd)
    try:
        win32_dlg.set_focus()
    except Exception:
        pass # 如果已經在最上層導致報錯，就不管它，繼續執行
    time.sleep(0.5)

    # 3. 鎖定樹狀圖
    try:
        tree = win32_dlg.child_window(class_name_re=".*SysTreeView32.*")
        tree.set_focus()
    except Exception:
        pass
    time.sleep(0.5)

    # 4. 鍵盤物理外掛 (拆解動作，放慢速度，確保老系統跟得上)
    try:
        send_keys("{HOME}")     # 回到最頂層「系統的」
        time.sleep(0.5)
        
        send_keys("{DOWN}")     # 第一下往下：來到「管理者」
        time.sleep(0.5)
        
        send_keys("{DOWN}")     # 第二下往下：來到「使用者」
        time.sleep(0.5)
        
        send_keys("{RIGHT}")    # 展開「使用者」
        time.sleep(1.5)         # 多給一點時間等待資料夾展開動畫
        
        send_keys("{DOWN}")     # 往下選取裡面的第一個報表 (00981A)
        time.sleep(0.5)

        # 5. 按下確定
        try:
            win32_dlg.child_window(title="確定(Y)", control_type="Button").click()
        except Exception:
            send_keys("%y")     # 備用快捷鍵 Alt+Y
            time.sleep(0.3)
            send_keys("{ENTER}")

    except Exception as e:
        raise RuntimeError(f"選取報表失敗: {e}")


def _read_sheet_data(ws) -> list:
    try:
        used = ws.UsedRange
        rows_out = []
        for r in range(1, used.Rows.Count + 1):
            row = []
            for c in range(1, used.Columns.Count + 1):
                val = ws.Cells(r, c).Value
                if val is None:
                    row.append("")
                elif isinstance(val, float):
                    row.append(str(int(val)) if val == int(val) else f"{val:.6g}")
                else:
                    row.append(str(val).strip())
            line = " ".join(v for v in row if v)
            if ETF in line:
                rows_out.append(line)
        return rows_out
    except Exception as e:
        _p(f"讀取 Excel 儲存格失敗: {e}", "WARN")
        return []


def _hwnd_by_class(cls: str):
    result = []
    def cb(hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            try:
                if win32gui.GetClassName(hwnd) == cls:
                    result.append(hwnd)
            except Exception:
                pass
    win32gui.EnumWindows(cb, None)
    return result[0] if result else None


# ══════════════════════════════════════════════════════
#  Step4: CMoney 小綠 → 基金規模
# ══════════════════════════════════════════════════════

def fetch_fund_scale(on_progress=None) -> tuple:
    p = lambda m, lv="INFO": _p(m, lv, on_progress)
    p("開啟 CMoney 小綠...")

    hwnd = _find_cmoney_hwnd()
    if not hwnd:
        subprocess.Popen([CMONEY_EXE])
        p("CMoney 小綠 啟動中，等待 15 秒...")
        time.sleep(15)
        hwnd = _find_cmoney_hwnd()
        if not hwnd:
            raise RuntimeError("CMoney 小綠 無法找到（請確認已安裝並可正常執行）")

    win = None
    for backend in ("uia", "win32"):
        try:
            app = Application(backend=backend).connect(handle=hwnd, timeout=5)
            win = app.window(handle=hwnd)
            win.set_focus()
            time.sleep(1)
            break
        except Exception:
            continue
    if win is None:
        raise RuntimeError("無法連接 CMoney 小綠視窗")

    p("搜尋 ETF折溢價表...")
    _cmoney_search(win, "ETF折溢價表")
    time.sleep(3)

    p("切換個股 → 輸入 00981A...")
    _cmoney_goto_individual(win, ETF)
    time.sleep(3)

    p("讀取 基金資產價值 欄位...")
    yscale, tscale = _cmoney_read_scale(win)
    p(f"基金規模：昨天={yscale}億  今天={tscale}億", "OK")
    return yscale, tscale


def _find_cmoney_hwnd():
    result = []
    def cb_title(hwnd, _):
        if not win32gui.IsWindowVisible(hwnd):
            return
        t = win32gui.GetWindowText(hwnd)
        cls = win32gui.GetClassName(hwnd)
        
        # 核心修正：嚴格排除 Excel 增益集視窗，防止小綠認錯人跑去亂打字！
        if "CMoneyExcel" in t or "資料轉出精靈" in t or "自訂報表" in t:
            return
            
        if any(k in t for k in ("CMoney", "理財寶", "法人", "小綠")):
            result.append(hwnd)
        elif cls == "ThunderRT6Main" and t:
            result.append(hwnd)

    win32gui.EnumWindows(cb_title, None)
    if result:
        result.sort(key=_window_area, reverse=True)
        return result[0]

    cmoney_exe = os.path.basename(CMONEY_EXE).lower()
    target_pids = set()
    try:
        import subprocess as _sp
        out = _sp.check_output(
            ["tasklist", "/FI", f"IMAGENAME eq {cmoney_exe}", "/FO", "CSV", "/NH"],
            text=True, errors="ignore"
        )
        for line in out.splitlines():
            parts = line.strip().strip('"').split('","')
            if len(parts) >= 2:
                try:
                    target_pids.add(int(parts[1]))
                except ValueError:
                    pass
    except Exception:
        pass

    if not target_pids:
        return None

    proc_result = []
    def cb_proc(hwnd, _):
        if not win32gui.IsWindowVisible(hwnd):
            return
        try:
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            if pid in target_pids and win32gui.GetWindowText(hwnd):
                proc_result.append(hwnd)
        except Exception:
            pass

    win32gui.EnumWindows(cb_proc, None)
    if proc_result:
        proc_result.sort(key=_window_area, reverse=True)
        return proc_result[0]
    return None


def _window_area(hwnd):
    try:
        r = win32gui.GetWindowRect(hwnd)
        return (r[2] - r[0]) * (r[3] - r[1])
    except Exception:
        return 0


def _cmoney_search(win, keyword: str):
    found = False
    try:
        edit = win.child_window(control_type="Edit", found_index=0)
        edit.click_input()
        time.sleep(0.2)
        edit.select()
        edit.type_keys(keyword, with_spaces=True)
        send_keys("{ENTER}")
        found = True
    except Exception:
        pass

    if not found:
        try:
            hwnd = win.handle
            app32 = Application(backend="win32").connect(handle=hwnd)
            w32 = app32.window(handle=hwnd)
            edit32 = w32.child_window(class_name_re=r"Thunder.*TextBox|Thunder.*Edit|Edit", found_index=0)
            edit32.click_input()
            time.sleep(0.2)
            edit32.select()
            edit32.type_keys(keyword, with_spaces=True)
            send_keys("{ENTER}")
            found = True
        except Exception as e:
            _p(f"win32 搜尋框失敗: {e}", "WARN")

    if not found:
        try:
            win.set_focus()
            send_keys("^f")
            time.sleep(0.5)
            send_keys(keyword + "{ENTER}", pause=0.05)
            found = True
            _p("CMoney 搜尋：Ctrl+F 備用", "WARN")
        except Exception as e:
            _p(f"CMoney 搜尋 Ctrl+F 失敗: {e}", "WARN")

    if not found:
        try:
            win.set_focus()
            time.sleep(0.3)
            send_keys(keyword + "{ENTER}", pause=0.05)
            _p("CMoney 搜尋：直接鍵入備用", "WARN")
        except Exception as e:
            _p(f"CMoney 搜尋全失敗: {e}", "WARN")


def _cmoney_goto_individual(win, code: str):
    for ctrl_type in ("TabItem", "RadioButton"):
        try:
            win.child_window(title="個股", control_type=ctrl_type).click_input()
            time.sleep(0.5)
            break
        except Exception:
            continue
    try:
        edits = win.descendants(control_type="Edit")
        for edit in reversed(edits):
            try:
                edit.click_input()
                edit.set_text(code)
                send_keys("{ENTER}")
                return
            except Exception:
                continue
    except Exception as e:
        _p(f"個股代號輸入失敗: {e}", "WARN")


def _cmoney_read_scale(win) -> tuple:
    texts = []
    try:
        for ctrl in win.descendants():
            try:
                t = ctrl.window_text().strip()
                if t:
                    texts.append(t)
            except Exception:
                pass
    except Exception:
        pass
    full = "\n".join(texts)

    candidates = re.findall(r"\b(\d{3,4}\.\d)\b", full)
    if len(candidates) >= 2:
        return candidates[-2], candidates[-1]
    elif len(candidates) == 1:
        return candidates[0], candidates[0]

    raise RuntimeError(
        "無法自動讀取基金規模。\n"
        "請在 CMoney 小綠 → ETF折溢價表 → 個股 → 00981A → 基金資產價值\n"
        "手動取得昨天和今天的數值（例：1162.4）並填入 config.json 的 manual_scale 欄位。"
    )


# ══════════════════════════════════════════════════════
#  Step5: AI Studio 分析
# ══════════════════════════════════════════════════════

def analyze_in_aistudio(excel_data: str, yscale: str, tscale: str, on_progress=None) -> dict:
    p = lambda m, lv="INFO": _p(m, lv, on_progress)
    DOWNLOAD_DIR.mkdir(exist_ok=True)

    opts = Options()
    opts.add_argument(f"--user-data-dir={PROFILE_DIR}")
    opts.add_argument("--start-maximized")
    opts.add_argument("--disable-notifications")
    opts.add_experimental_option("excludeSwitches", ["enable-automation"])
    opts.add_experimental_option("prefs", {
        "download.default_directory": str(DOWNLOAD_DIR),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
    })

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=opts)
    wait = WebDriverWait(driver, 60)

    try:
        p("開啟 AI Studio ETF AlphaTracker...")
        driver.get(AI_STUDIO_URL)
        time.sleep(6)

        for _ in range(24):
            if "aistudio.google.com" in driver.current_url:
                break
            p("等待 Google 登入（請在瀏覽器完成）...")
            time.sleep(5)

        p("點擊「智慧貼上（自動合併）」...")
        try:
            btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[contains(text(),'智慧貼上') or contains(text(),'自動合併')]")))
            btn.click()
            time.sleep(1)
        except TimeoutException:
            p("找不到智慧貼上按鈕，繼續嘗試...", "WARN")

        p("貼入 Excel 持股資料（剪貼簿）...")
        pyperclip.copy(excel_data)
        try:
            ta = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "textarea")))
            ta.click()
            ta.send_keys(Keys.CONTROL, "a")
            ta.send_keys(Keys.CONTROL, "v")
        except TimeoutException:
            try:
                ta = driver.find_element(By.CSS_SELECTOR, "[contenteditable='true']")
                ta.click()
                ta.send_keys(Keys.CONTROL, "a")
                ta.send_keys(Keys.CONTROL, "v")
            except Exception as e:
                p(f"貼上失敗: {e}", "WARN")
        time.sleep(1)

        p(f"填入基金規模：昨天={yscale}  今天={tscale}...")
        try:
            inputs = [i for i in driver.find_elements(By.CSS_SELECTOR, "input[type='number'], input[type='text']") if i.is_displayed()]
            if len(inputs) >= 2:
                inputs[-2].clear(); inputs[-2].send_keys(yscale)
                inputs[-1].clear(); inputs[-1].send_keys(tscale)
            elif len(inputs) == 1:
                inputs[0].clear(); inputs[0].send_keys(f"{yscale} {tscale}")
        except Exception as e:
            p(f"填入規模失敗: {e}", "WARN")

        p("點擊「開始分析」...")
        try:
            btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[contains(text(),'開始分析')]")))
            btn.click()
        except TimeoutException as e:
            p(f"開始分析按鈕找不到: {e}", "WARN")

        p("等待 AI Studio 分析完成（最多 90 秒）...")
        result = {"建倉": [], "加碼": [], "減碼": [], "清倉": []}
        for _ in range(30):
            time.sleep(3)
            body = driver.find_element(By.TAG_NAME, "body").text
            if "成分股變動" in body or "加碼" in body or "建倉" in body:
                p("分析完成，開始解析結果...")
                result = _parse_holdings(body)
                break

        p("下載分析報告圖表...")
        chart_path = _download_chart(driver, wait)
        result["chart_path"] = chart_path

        total = sum(len(v) for v in result.values() if isinstance(v, list))
        p(f"AI Studio 完成：共 {total} 筆持股變動", "OK")
        return result

    finally:
        time.sleep(2)
        driver.quit()


def _parse_holdings(text: str) -> dict:
    建倉, 加碼, 減碼, 清倉 = [], [], [], []
    lines = [l.strip() for l in text.splitlines() if l.strip()]

    i = 0
    while i < len(lines):
        line = lines[i]
        if re.match(r"^\d{4,5}$", line) and i + 3 < len(lines):
            code       = line
            name       = lines[i + 1]
            action_raw = lines[i + 2]
            change_raw = lines[i + 3]

            action = ""
            if any(k in action_raw for k in ("建倉", "新增")):
                action = "建倉"
            elif any(k in action_raw for k in ("加碼", "加")):
                action = "加碼"
            elif any(k in action_raw for k in ("清倉", "刪除")):
                action = "清倉"
            elif any(k in action_raw for k in ("減碼", "減")):
                action = "減碼"

            m = re.search(r"([+-]?[\d,]+)", change_raw)
            shares = abs(int(m.group(1).replace(",", ""))) if m else 0

            if action:
                label = f"{name}({code})"
                if action in ("建倉", "加碼"):
                    (建倉 if action == "建倉" else 加碼).append(f"{label}+{shares}張")
                else:
                    (清倉 if action == "清倉" else 減碼).append(f"{label}-{shares}張")
            i += 4
            continue
        i += 1

    return {"建倉": 建倉, "加碼": 加碼, "減碼": 減碼, "清倉": 清倉}


def _download_chart(driver, wait) -> str:
    before = set(DOWNLOAD_DIR.glob("*"))
    try:
        btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[contains(text(),'下載分析報告圖表') or contains(text(),'下載圖表') or contains(text(),'下載')]")))
        btn.click()
        for _ in range(20):
            time.sleep(1)
            after = set(DOWNLOAD_DIR.glob("*"))
            new_files = after - before
            if new_files:
                newest = max(new_files, key=lambda f: f.stat().st_mtime)
                return str(newest)
    except Exception as e:
        _p(f"圖表下載失敗: {e}", "WARN")
    return ""


def run_all(on_progress=None) -> dict:
    _p("=== 全自動資料抓取開始 ===", "INFO", on_progress)
    excel_data = fetch_holdings_from_excel(on_progress)
    yscale, tscale = fetch_fund_scale(on_progress)
    result = analyze_in_aistudio(excel_data, yscale, tscale, on_progress)
    _p("=== 資料抓取完成 ===", "OK", on_progress)
    return result
