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
    """
    Opens Excel, uses CMoney VSTO add-in (00981A操作日報 template),
    pulls holdings data, reads all rows, returns as space-separated text.
    """
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
    win.maximize() # 強制最大化，確保 UI 元素不被隱藏折疊
    time.sleep(1)

    # 1. 點擊 增益集 標籤 (對應圖01的標示1)
    p("點擊 增益集 標籤...")
    try:
        win.child_window(title="增益集", control_type="TabItem").click_input()
        time.sleep(1)
    except Exception as e:
        p(f"UIA 增益集標籤點擊失敗: {e}，請確認是否已載入增益集", "WARN")
        raise

    # 2. 點擊 自訂工具列 的第一個按鈕 (對應圖01的標示2)
    p("點擊 自訂工具列 按鈕...")
    try:
        # 直接鎖定「自訂工具列」群組，並點擊其底下的第一個按鈕，完全棄用座標點擊
        grp = win.child_window(title="自訂工具列", control_type="Group")
        btns = grp.children(control_type="Button")
        if btns:
            btns[0].click_input()
            p("成功點擊增益集按鈕", "INFO")
        else:
            raise RuntimeError("自訂工具列中沒有找到按鈕")
    except Exception as e:
        raise RuntimeError(f"增益集按鈕點擊失敗，流程中斷：{e}")

    time.sleep(2)
    _handle_cmoney_dialog(on_progress)


def _handle_cmoney_dialog(on_progress=None):
    p = lambda m, lv="INFO": _p(m, lv, on_progress)

    dlg = None
    for _ in range(15):
        try:
            # 根據截圖，精準匹配標題
            dlg = Desktop(backend="uia").window(title_re=r".*CMoneyExcel.*|.*資料匯出.*|.*資料轉出.*")
            if dlg.exists():
                break
        except Exception:
            pass
        time.sleep(1)
    if not dlg or not dlg.exists():
        raise RuntimeError("CMoneyExcel 對話框未出現")

    dlg.set_focus()
    time.sleep(0.5)

    # ── Step A: 型式選擇 → 下一步 (對應圖02的標示3) ──
    try:
        next_btn = dlg.child_window(title_re=r"下一步.*", control_type="Button")
        if next_btn.exists(timeout=2):
            p("型式選擇 → 點擊 下一步>...")
            next_btn.click_input()
            time.sleep(2)
    except Exception:
        pass  # 找不到下一步按鈕代表可能已經直接進入主畫面

    # ── Step B: 主對話框 → 開啟... (對應圖03的標示4) ──
    p("點擊 開啟...")
    try:
        dlg.child_window(title_re=r"開啟.*", control_type="Button").click_input()
    except Exception as e:
        raise RuntimeError(f"找不到「開啟...」按鈕: {e}")
    time.sleep(1.5)

    # ── Step C: 開啟自訂報表 → 展開使用者 → 00981A 操作日報 (對應圖04、05的標示5、6) ──
    p("選擇 00981A 操作日報...")
    _select_cmoney_template()
    time.sleep(1)

    # ── Step D: 主對話框 → 確定 ──
    p("主對話框 → 確定...")
    try:
        dlg = Desktop(backend="uia").window(title_re=r".*CMoneyExcel.*|.*資料匯出.*|.*資料轉出.*")
        dlg.set_focus()
        time.sleep(0.3)
        dlg.child_window(title="確定", control_type="Button").click_input()
    except Exception:
        send_keys("{ENTER}")
    
    # 移除原本強制去點綠色更新箭頭的邏輯（_click_excel_refresh），
    # 因為點擊確定後，CMoney 增益集通常就會自動拋轉資料。
    time.sleep(5)


def _select_cmoney_template():
    """開啟自訂報表對話框：展開「使用者」→ 點選「00981A 操作日報」→ 確定"""
    try:
        # Dialog title = "開啟自訂報表"
        dlg = Desktop(backend="uia").window(title_re=r".*開啟自訂報表.*|.*自訂報表.*")
        dlg.set_focus()
        time.sleep(0.5)

        # Find tree and expand 使用者 node
        tree = dlg.child_window(control_type="Tree")
        user_node = None
        for item in tree.children():
            try:
                if "使用者" in item.window_text():
                    user_node = item
                    break
            except Exception:
                continue

        if user_node:
            try:
                user_node.expand()
            except Exception:
                user_node.click_input()
            time.sleep(0.5)

            # Click 00981A 操作日報 under 使用者
            for child in user_node.children():
                try:
                    if "00981A" in child.window_text():
                        child.click_input()
                        break
                except Exception:
                    continue
        else:
            # Fallback: search entire dialog for 00981A item
            try:
                dlg.child_window(title_re=r".*00981A.*").click_input()
            except Exception:
                pass

        time.sleep(0.3)
        dlg.child_window(title="確定", control_type="Button").click_input()

    except Exception as e:
        _p(f"開啟自訂報表失敗（繼續）: {e}", "WARN")
        send_keys("{ENTER}")



def _click_excel_refresh():
    hwnd = _hwnd_by_class("XLMAIN")
    if not hwnd:
        return
    try:
        app = Application(backend="uia").connect(handle=hwnd)
        win = app.window(handle=hwnd)
        win.set_focus()
        time.sleep(0.3)
        # Try known refresh button titles
        for title in ["更新", "重新整理", "Refresh", "↑", "Update"]:
            try:
                win.child_window(title=title, control_type="Button").click_input()
                return
            except Exception:
                continue
        # Fallback: second button in ribbon (first is 新增自訂表格)
        btns = [b for b in win.descendants(control_type="Button")
                if b.rectangle().top < 180]
        if len(btns) >= 2:
            btns[1].click_input()
    except Exception as e:
        _p(f"綠色箭頭點擊失敗（可能不需要）: {e}", "WARN")


def _read_sheet_data(ws) -> list:
    """Read data rows from Excel sheet via win32com. Returns list of space-joined rows."""
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
    """
    Opens CMoney 小綠, searches ETF折溢價表 for 00981A,
    reads 基金資產價值 for last 2 trading days.
    Returns (yesterday_scale, today_scale) e.g. ("1162.4", "1165.2")
    """
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

    # Try uia first, then win32
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
    """Find CMoney 小綠 main window handle.
    Searches by title first, then falls back to process-based detection."""
    result = []

    def cb_title(hwnd, _):
        if not win32gui.IsWindowVisible(hwnd):
            return
        t = win32gui.GetWindowText(hwnd)
        cls = win32gui.GetClassName(hwnd)
        # Match by known title keywords OR by CMoney's VB6 main form class
        if any(k in t for k in ("CMoney", "理財寶", "法人", "小綠")):
            result.append(hwnd)
        elif cls == "ThunderRT6Main" and t:
            result.append(hwnd)

    win32gui.EnumWindows(cb_title, None)
    if result:
        # Prefer the largest window (main UI, not floating toolbar)
        result.sort(key=_window_area, reverse=True)
        return result[0]

    # Fallback: find windows belonging to CMoney.exe process
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
    """Type keyword into CMoney search box. ThunderRT6Main uses custom controls."""
    found = False

    # Strategy 1: UIA Edit
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

    # Strategy 2: win32 backend Edit (ThunderRT6FormDC / ThunderRT6TextBox)
    if not found:
        try:
            hwnd = win.handle
            app32 = Application(backend="win32").connect(handle=hwnd)
            w32 = app32.window(handle=hwnd)
            edit32 = w32.child_window(class_name_re=r"Thunder.*TextBox|Thunder.*Edit|Edit",
                                      found_index=0)
            edit32.click_input()
            time.sleep(0.2)
            edit32.select()
            edit32.type_keys(keyword, with_spaces=True)
            send_keys("{ENTER}")
            found = True
        except Exception as e:
            _p(f"win32 搜尋框失敗: {e}", "WARN")

    # Strategy 3: Ctrl+F shortcut then type
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

    # Strategy 4: just type (window may have focus on search by default)
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
    """Extract 基金資產價值 values (###.#) from CMoney grid text."""
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

    # 基金規模 in 億: values like 1162.4, 483.0, 1233.5
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

def analyze_in_aistudio(excel_data: str, yscale: str, tscale: str,
                         on_progress=None) -> dict:
    """
    Pastes data into AI Studio ETF AlphaTracker, runs analysis,
    parses 建倉/加碼/減碼/清倉, downloads chart.
    Returns {建倉, 加碼, 減碼, 清倉, chart_path}.
    """
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

    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()), options=opts
    )
    wait = WebDriverWait(driver, 60)

    try:
        p("開啟 AI Studio ETF AlphaTracker...")
        driver.get(AI_STUDIO_URL)
        time.sleep(6)

        # Wait for Google login if redirected
        for _ in range(24):
            if "aistudio.google.com" in driver.current_url:
                break
            p("等待 Google 登入（請在瀏覽器完成）...")
            time.sleep(5)

        p("點擊「智慧貼上（自動合併）」...")
        try:
            btn = wait.until(EC.element_to_be_clickable(
                (By.XPATH,
                 "//*[contains(text(),'智慧貼上') or contains(text(),'自動合併')]")
            ))
            btn.click()
            time.sleep(1)
        except TimeoutException:
            p("找不到智慧貼上按鈕，繼續嘗試...", "WARN")

        p("貼入 Excel 持股資料（剪貼簿）...")
        pyperclip.copy(excel_data)
        try:
            ta = wait.until(EC.presence_of_element_located(
                (By.CSS_SELECTOR, "textarea")
            ))
            ta.click()
            ta.send_keys(Keys.CONTROL, "a")
            ta.send_keys(Keys.CONTROL, "v")
        except TimeoutException:
            try:
                ta = driver.find_element(
                    By.CSS_SELECTOR, "[contenteditable='true']"
                )
                ta.click()
                ta.send_keys(Keys.CONTROL, "a")
                ta.send_keys(Keys.CONTROL, "v")
            except Exception as e:
                p(f"貼上失敗: {e}", "WARN")
        time.sleep(1)

        p(f"填入基金規模：昨天={yscale}  今天={tscale}...")
        try:
            inputs = [i for i in driver.find_elements(
                By.CSS_SELECTOR, "input[type='number'], input[type='text']"
            ) if i.is_displayed()]
            if len(inputs) >= 2:
                inputs[-2].clear(); inputs[-2].send_keys(yscale)
                inputs[-1].clear(); inputs[-1].send_keys(tscale)
            elif len(inputs) == 1:
                inputs[0].clear(); inputs[0].send_keys(f"{yscale} {tscale}")
        except Exception as e:
            p(f"填入規模失敗: {e}", "WARN")

        p("點擊「開始分析」...")
        try:
            btn = wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//*[contains(text(),'開始分析')]")
            ))
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
    """Parse 建倉/加碼/減碼/清倉 from AI Studio body text."""
    建倉, 加碼, 減碼, 清倉 = [], [], [], []
    lines = [l.strip() for l in text.splitlines() if l.strip()]

    i = 0
    while i < len(lines):
        line = lines[i]
        # Stock code line: 4-5 digit number alone on a line
        if re.match(r"^\d{4,5}$", line) and i + 3 < len(lines):
            code       = line
            name       = lines[i + 1]
            action_raw = lines[i + 2]   # e.g. "↑ 加碼" or "加碼"
            change_raw = lines[i + 3]   # e.g. "+30,000" or "+30,000股"

            # Determine action category
            action = ""
            if any(k in action_raw for k in ("建倉", "新增")):
                action = "建倉"
            elif any(k in action_raw for k in ("加碼", "加")):
                action = "加碼"
            elif any(k in action_raw for k in ("清倉", "刪除")):
                action = "清倉"
            elif any(k in action_raw for k in ("減碼", "減")):
                action = "減碼"

            # Extract share count
            m = re.search(r"([+-]?[\d,]+)", change_raw)
            shares = abs(int(m.group(1).replace(",", ""))) if m else 0

            if action:
                label = f"{name}({code})"
                if action in ("建倉", "加碼"):
                    (建倉 if action == "建倉" else 加碼).append(
                        f"{label}+{shares}張"
                    )
                else:
                    (清倉 if action == "清倉" else 減碼).append(
                        f"{label}-{shares}張"
                    )
            i += 4
            continue
        i += 1

    return {"建倉": 建倉, "加碼": 加碼, "減碼": 減碼, "清倉": 清倉}


def _download_chart(driver, wait) -> str:
    before = set(DOWNLOAD_DIR.glob("*"))
    try:
        btn = wait.until(EC.element_to_be_clickable(
            (By.XPATH,
             "//*[contains(text(),'下載分析報告圖表') or "
             "contains(text(),'下載圖表') or contains(text(),'下載')]")
        ))
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


# ══════════════════════════════════════════════════════
#  Main entry
# ══════════════════════════════════════════════════════

def run_all(on_progress=None) -> dict:
    """Run complete data-fetch pipeline. Returns dict for poster."""
    _p("=== 全自動資料抓取開始 ===", "INFO", on_progress)
    excel_data = fetch_holdings_from_excel(on_progress)
    yscale, tscale = fetch_fund_scale(on_progress)
    result = analyze_in_aistudio(excel_data, yscale, tscale, on_progress)
    _p("=== 資料抓取完成 ===", "OK", on_progress)
    return result
