import tkinter as tk
from tkinter import font as tkfont
import customtkinter as ctk
import time, re, json, os
import ctypes, ctypes.wintypes
from datetime import datetime, timedelta, timezone

# --- 高 DPI 支援（讓工作列圖標清晰）---
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)  # Per-monitor DPI aware
except:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except:
        pass

# --- 設定 CTK 基本外觀 ---
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

# --- 設定儲存路徑 ---
appdata_dir = os.path.join(os.environ['LOCALAPPDATA'], 'Floating-Clock')
if not os.path.exists(appdata_dir):
    try: os.makedirs(appdata_dir)
    except: appdata_dir = "."
CONFIG_FILE = os.path.join(appdata_dir, "clock_config.json")

# 預設設定 (完全不動)
DEFAULT_CONFIG = {
    "bg_alpha": 0.8, "text_color": "#FFFFFF", "bg_color": "#1A1A1A",
    "is_24h": True, "show_ampm": True, "show_weekday": False,
    "weekday_lang": "ZH", "show_timezone": False, "utc_offset": 8,
    "show_timestamp": False, "show_ms": False,
    "excel_date_format": "YYYY-MM-DD hh:mm:ss",
    "is_topmost": True,
    "font_size": 25,
    "language": "ZH",
    "clock_font_family": "Microsoft JhengHei",
    "appearance_mode": "dark",
    "text_alpha": 1.0,
    "bg_fill_alpha": 0.8,
    "recent_colors": [],
}

# 中英文字對照
LANG_TEXT = {
    "ZH": {
        "msg_ok": "確定",
        "msg_cancel": "取消",
        "msg_hint": "提示",
        "copy_display": "複製顯示內容",
        "copy_ts": "複製當前時間戳",
        "top_on": "取消置頂",
        "top_off": "設為置頂",
        "fmt_12h": "切換為 12H",
        "fmt_24h": "切換為 24H",
        "hide_ampm": "隱藏 AM/PM",
        "show_ampm": "顯示 AM/PM",
        "hide_weekday": "隱藏星期",
        "show_weekday": "顯示星期",
        "week_en": "英文展示星期",
        "week_zh": "中文展示星期",
        "hide_tz": "隱藏時區",
        "show_tz": "顯示時區",
        "change_tz": "切換時區",
        "reset_sys_tz": "重設為系統時區",
        "custom_fmt": "自定義時間格式",
        "hide_ts": "隱藏時間戳",
        "show_ts": "顯示時間戳",
        "show_sec": "秒展示",
        "show_ms": "毫秒展示",
        "ts_converter": "時間戳轉換器",
        "change_fg": "變更文字顏色",
        "change_bg": "變更背景顏色",
        "use_dark_theme": "使用深色",
        "use_light_theme": "使用淺色",
        "reset_style": "恢復預設樣式",
        "exit": "結束時鐘",
        "converter_title": "時間戳轉換器",
        "converter_label_dt": "日期時間 (YYYY-MM-DD hh:mm:ss.SSS)",
        "converter_label_ts": "時間戳 (ms/s)",
        "btn_fetch_now": "抓取目前時間",
        "btn_start_conv": "開始轉換",
        "fmt_dialog_title": "自定義時間格式",
        "tz_dialog_title": "切換時區",
        "tz_dialog_prompt": "滾輪每次調整 15 分鐘。\n也可輸入：+8、+8.25、+08:15（範圍 UTC-12:00 ~ UTC+14:00）",
        "tz_dialog_current": "目前時區",
        "tz_dialog_input": "輸入時區偏移",
        "pick_fg_title": "變更文字顏色",
        "pick_bg_title": "變更背景顏色",
        "err_ts_len": "時間戳請輸入 10 碼(秒)或 13 碼(毫秒)",
        "err_format": "格式輸入錯誤",
        "lang_to_en": "English Interface",
        "lang_to_zh": "中文介面",
        "help": "說明",
    },
    "EN": {
        "msg_ok": "OK",
        "msg_cancel": "Cancel",
        "msg_hint": "Hint",
        "copy_display": "Copy display text",
        "copy_ts": "Copy current timestamp",
        "top_on": "Unpin (cancel always on top)",
        "top_off": "Pin (set always on top)",
        "fmt_12h": "Switch to 12H",
        "fmt_24h": "Switch to 24H",
        "hide_ampm": "Hide AM/PM",
        "show_ampm": "Show AM/PM",
        "hide_weekday": "Hide weekday",
        "show_weekday": "Show weekday",
        "week_en": "Show weekday: English",
        "week_zh": "Show weekday: Chinese",
        "hide_tz": "Hide timezone",
        "show_tz": "Show timezone",
        "change_tz": "Change timezone",
        "reset_sys_tz": "Reset to system timezone",
        "custom_fmt": "Custom time format",
        "hide_ts": "Hide timestamp",
        "show_ts": "Show timestamp",
        "show_sec": "Show seconds",
        "show_ms": "Show milliseconds",
        "ts_converter": "Timestamp converter",
        "change_fg": "Change text color",
        "change_bg": "Change background color",
        "use_dark_theme": "Use dark theme",
        "use_light_theme": "Use light theme",
        "reset_style": "Reset to default style",
        "exit": "End clock",
        "converter_title": "Timestamp converter",
        "converter_label_dt": "Datetime (YYYY-MM-DD hh:mm:ss.SSS)",
        "converter_label_ts": "Timestamp (ms/s)",
        "btn_fetch_now": "Use current time",
        "btn_start_conv": "Convert",
        "fmt_dialog_title": "Custom time format",
        "tz_dialog_title": "Change timezone",
        "tz_dialog_prompt": "Mouse wheel adjusts by 15 minutes.\nYou can also type: +8, +8.25, +08:15 (Range: UTC-12:00 to UTC+14:00)",
        "tz_dialog_current": "Current",
        "tz_dialog_input": "Input",
        "pick_fg_title": "Change text color",
        "pick_bg_title": "Change background color",
        "err_ts_len": "Timestamp must be 10 digits (sec) or 13 digits (ms)",
        "err_format": "Invalid format",
        "lang_to_en": "English Interface",
        "lang_to_zh": "中文介面",
        "help": "Help",
    },
}

class FloatingClock:
    def __init__(self, root):
        self.root = root
        self.root.title("Floating-Clock")
        self.root.overrideredirect(True)
        self.transparent_key = "#010203"
        self.root.configure(bg=self.transparent_key)
        try:
            self.root.attributes("-transparentcolor", self.transparent_key)
        except:
            pass

        # 載入應用程式圖標（打包後從同目錄或資源目錄找）
        self._app_icon = None
        self._set_app_icon()

        self.load_config()
        self.apply_appearance_mode()
        self.root.attributes("-topmost", self.is_topmost)

        sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
        self.root.geometry(f"+{sw//2 - 150}+{sh//2 - 50}")

        # ── 啟動骨架：在時鐘 label 出現前顯示旋轉動畫 ──
        skel_bg = self.bg_color
        skel_w, skel_h = 300, 70
        self._skel_canvas = tk.Canvas(root, width=skel_w, height=skel_h,
                                      bg=skel_bg, highlightthickness=0)
        self._skel_canvas.pack(padx=25, pady=15)
        spin_cx, spin_cy, spin_r = skel_w // 2, skel_h // 2, 16
        arc_color = "#3A7EBF" if self.appearance_mode == "dark" else "#4A90D9"
        track_color = "#444444" if self.appearance_mode == "dark" else "#CCCCCC"
        # 軌道圓環
        self._skel_canvas.create_oval(
            spin_cx - spin_r, spin_cy - spin_r,
            spin_cx + spin_r, spin_cy + spin_r,
            outline=track_color, width=3)
        self._spin_arc = self._skel_canvas.create_arc(
            spin_cx - spin_r, spin_cy - spin_r,
            spin_cx + spin_r, spin_cy + spin_r,
            start=90, extent=80, outline=arc_color, width=3, style="arc")
        self._spin_angle = [0]

        def _spin():
            if self._skel_canvas is None:
                return
            try:
                if not self._skel_canvas.winfo_exists():
                    return
            except:
                return
            self._spin_angle[0] = (self._spin_angle[0] - 12) % 360
            self._skel_canvas.itemconfig(
                self._spin_arc, start=self._spin_angle[0])
            self._spin_after = root.after(40, _spin)

        self._spin_after = root.after(40, _spin)

        # 時鐘字體加粗
        self.main_font = (self.clock_font_family, self.font_size, 'bold')
        self.label = tk.Label(root, font=self.main_font,
                              fg=self.text_color, bg=self.bg_color, padx=25, pady=15, justify='center')
        # label 先不 pack，等第一次 update_clock 完成後才替換骨架

        self.label.bind("<Button-1>", self.start_move)
        self.label.bind("<B1-Motion>", self.do_move)
        self.label.bind("<Double-Button-1>", lambda e: self.safe_exit())
        self.label.bind("<Button-3>", self.open_ctk_main_menu)

        self.root.bind("<MouseWheel>", self.on_mouse_wheel)
        self.root.bind("<Control-MouseWheel>", self.on_font_zoom)
        self.root.bind_all("<Control-Shift-C>", lambda e: self.copy_smart_timestamp())

        self.update_clock()

    # 取目前語系文字
    def t(self, key):
        lang = getattr(self, "language", "ZH")
        return LANG_TEXT.get(lang, LANG_TEXT["ZH"]).get(key, key)

    def apply_appearance_mode(self):
        mode = getattr(self, "appearance_mode", "dark")
        if mode not in ("dark", "light"):
            mode = "dark"
        self.appearance_mode = mode
        ctk.set_appearance_mode(mode)
        # 重繪時鐘 label 顏色
        if hasattr(self, "label"):
            self.apply_alpha_settings()

    def refresh_clock_font(self):
        self.main_font = (self.clock_font_family, self.font_size, "bold")
        if hasattr(self, "label"):
            self.label.config(font=self.main_font)

    def hex_to_rgb(self, color_hex):
        s = color_hex.lstrip("#")
        if len(s) != 6:
            return (255, 255, 255)
        return (int(s[0:2], 16), int(s[2:4], 16), int(s[4:6], 16))

    def blend_hex(self, fg_hex, bg_hex, alpha):
        a = max(0.0, min(1.0, float(alpha)))
        fr, fg, fb = self.hex_to_rgb(fg_hex)
        br, bg, bb = self.hex_to_rgb(bg_hex)
        r = int(round(fr * a + br * (1 - a)))
        g = int(round(fg * a + bg * (1 - a)))
        b = int(round(fb * a + bb * (1 - a)))
        return f"#{r:02X}{g:02X}{b:02X}"

    def apply_alpha_settings(self):
        self.bg_alpha = max(0.1, min(1.0, float(getattr(self, "bg_alpha", 0.8))))
        self.text_alpha = max(0.1, min(1.0, float(getattr(self, "text_alpha", 1.0))))
        self.bg_fill_alpha = max(0.1, min(1.0, float(getattr(self, "bg_fill_alpha", 0.8))))
        effective_fg = self.blend_hex(self.text_color, self.bg_color, self.text_alpha)

        # dirty-flag：值沒有變就跳過 Tk 呼叫，避免每秒無謂觸發 -alpha / label.config
        new_state = (self.bg_fill_alpha, effective_fg, self.bg_color)
        if getattr(self, "_alpha_state_cache", None) == new_state:
            return
        self._alpha_state_cache = new_state

        self.root.attributes("-alpha", self.bg_fill_alpha)
        if hasattr(self, "label"):
            self.label.config(fg=effective_fg, bg=self.bg_color)


    def get_menu_colors(self):
        if getattr(self, "appearance_mode", "dark") == "light":
            return {
                "bg": "#EBEBEB",
                "fg": "#1F1F1F",
                "active_bg": "#D7E7FF",
                "active_fg": "#000000",
            }
        return {
            "bg": "#2B2B2B",
            "fg": "#EDEDED",
            "active_bg": "#3A7EBF",
            "active_fg": "#FFFFFF",
        }

    def calc_ctk_menu_width(self, labels, font_tuple=("Microsoft JhengHei", 12), extra_px=10, min_px=170, max_px=360):
        try:
            # 快取 Font 物件，避免每次開選單都重建
            cache_key = font_tuple
            if not hasattr(self, "_menu_font_cache"):
                self._menu_font_cache = {}
            if cache_key not in self._menu_font_cache:
                self._menu_font_cache[cache_key] = tkfont.Font(
                    root=self.root, family=font_tuple[0], size=font_tuple[1])
            f = self._menu_font_cache[cache_key]
            max_text = max((f.measure(str(t)) for t in labels), default=min_px)
            raw = max_text + extra_px
            screen_third = self.root.winfo_screenwidth() // 3
            if raw > screen_third:
                return min(max_px, max_text + 10)
            return max(min_px, raw)
        except:
            return min_px

    def clamp_utc_offset(self, val):
        try:
            v = float(val)
        except:
            v = 8.0
        # 15 分鐘一格
        v = round(v * 4) / 4
        return max(-12.0, min(14.0, v))

    def format_utc_offset(self, val=None):
        v = self.utc_offset if val is None else float(val)
        v = self.clamp_utc_offset(v)
        sign = "+" if v >= 0 else "-"
        abs_v = abs(v)
        h = int(abs_v)
        m = int(round((abs_v - h) * 60))
        if m == 0:
            return f"UTC{sign}{h}"
        return f"UTC{sign}{h}:{m:02d}"

    def parse_utc_offset_input(self, s):
        """
        支援:
        - +8 / -5
        - +8.25（小數小時）
        - +08:15 / -03:30
        回傳 float 小時（會被 clamp_utc_offset 正規化到 15 分鐘刻度與範圍）
        """
        if s is None:
            raise ValueError("empty")
        raw = str(s).strip()
        if not raw:
            raise ValueError("empty")

        raw = raw.upper().replace("UTC", "").strip()
        m = re.fullmatch(r"([+-]?)\s*(\d{1,2})(?::(\d{1,2}))?", raw)
        if m:
            sign = -1 if m.group(1) == "-" else 1
            hh = int(m.group(2))
            mm = int(m.group(3) or "0")
            if mm not in (0, 15, 30, 45):
                # 仍允許輸入其他分鐘，但會被四捨五入到 15 分鐘刻度
                pass
            return self.clamp_utc_offset(sign * (hh + mm / 60))

        # 小數小時格式
        v = float(raw)
        return self.clamp_utc_offset(v)

    # --- 取得定位 (彈窗出現在時鐘下方，下方不夠時往上) ---
    def get_popup_pos(self, popup_h=200):
        clock_x = self.root.winfo_x()
        clock_y = self.root.winfo_y()
        clock_h = self.root.winfo_height()
        sh = self.root.winfo_screenheight()
        below_y = clock_y + clock_h + 10
        above_y = clock_y - popup_h - 10
        if below_y + popup_h > sh:
            y = max(0, above_y)
        else:
            y = below_y
        return f"+{clock_x}+{y}"

    # --- 自定義 CTK 訊息視窗（單例，重複呼叫更新內容）---
    def show_ctk_message(self, title, message):
        # 若已有提示窗存在就直接更新，不重複開啟
        if hasattr(self, "_msg_win") and self._msg_win:
            try:
                if self._msg_win.winfo_exists():
                    self._msg_win.title(title)
                    self._msg_label.configure(text=message)
                    self._msg_win.lift()
                    self._msg_win.focus_force()
                    return
            except: pass

        msg_win = ctk.CTkToplevel(self.root)
        self.set_win_icon(msg_win)
        msg_win.title(title)
        msg_win.attributes("-topmost", True)
        msg_win.geometry(f"300x150{self.get_popup_pos(150)}")
        msg_win.resizable(False, False)
        msg_label = ctk.CTkLabel(msg_win, text=message, font=('Microsoft JhengHei', 12))
        msg_label.pack(pady=30)
        ctk.CTkButton(msg_win, text=self.t("msg_ok"), width=80,
                      font=('Microsoft JhengHei', 12),
                      command=msg_win.destroy).pack(pady=10)
        msg_win.lift()
        msg_win.focus_force()

        self._msg_win = msg_win
        self._msg_label = msg_label

        def _on_close():
            self._msg_win = None
        msg_win.bind("<Destroy>", lambda e: _on_close() if e.widget is msg_win else None)

    # --- 自定義輸入彈窗 (還原消失的提示文字與邏輯) ---
    def get_system_utc_offset(self):
        now = time.time()
        offset = (datetime.fromtimestamp(now) - datetime.utcfromtimestamp(now)).total_seconds()
        # 以 15 分鐘為一格（支援 :15/:30/:45）
        minutes = offset / 60
        minutes = int(round(minutes / 15) * 15)
        return self.clamp_utc_offset(minutes / 60)

    def load_config(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    saved = json.load(f)
                    for k, v in DEFAULT_CONFIG.items(): setattr(self, k, saved.get(k, v))
                    # 舊版可能是整數時區：這裡做一次正規化
                    self.utc_offset = self.clamp_utc_offset(getattr(self, "utc_offset", 8))
                    if "bg_fill_alpha" not in saved:
                        # 兼容舊版 text_alpha 欄位
                        self.bg_fill_alpha = float(saved.get("text_alpha", getattr(self, "bg_fill_alpha", 1.0)))
                    self.bg_alpha = float(saved.get("bg_alpha", getattr(self, "bg_alpha", 0.9)))
            except: self.apply_default()
        else: self.apply_default()

    def save_config(self):
        config_to_save = {k: getattr(self, k) for k in DEFAULT_CONFIG.keys()}
        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f: json.dump(config_to_save, f, indent=4)
        except: pass

    def safe_exit(self):
        self.save_config()
        self.root.quit()

    def apply_default(self):
        _preserve = {"recent_colors": getattr(self, "recent_colors", [])}
        for k, v in DEFAULT_CONFIG.items(): setattr(self, k, v)
        for k, v in _preserve.items(): setattr(self, k, v)
        self.apply_appearance_mode()
        self.refresh_clock_font()
        if hasattr(self, 'root'):
            self.root.attributes("-topmost", self.is_topmost)
            self.apply_alpha_settings()

    def copy_smart_timestamp(self):
        val = self.current_ts_val if (self.show_ms or not self.show_timestamp) else self.current_sec_val
        self.root.clipboard_clear()
        self.root.clipboard_append(val)

    def excel_to_py_format(self, fmt):
        mapping = {"YYYY": "%Y", "YY": "%y", "MM": "%m", "DD": "%d", "hh": "%H", "mm": "%M", "ss": "%S"}
        for k, v in mapping.items(): fmt = fmt.replace(k, v)
        return fmt

    def update_clock(self):
        # 第一次執行時：拆掉骨架，把真正的 label pack 上去
        if getattr(self, "_skel_canvas", None) is not None:
            try:
                if hasattr(self, "_spin_after"):
                    self.root.after_cancel(self._spin_after)
            except: pass
            try:
                self._skel_canvas.destroy()
            except: pass
            self._skel_canvas = None
            self.label.pack()
            self.apply_alpha_settings()

        self.utc_offset = self.clamp_utc_offset(self.utc_offset)
        tz = timezone(timedelta(hours=self.utc_offset))
        now = datetime.now(tz)
        raw_fmt = self.excel_date_format
        ms_str = str(now.microsecond // 1000).zfill(3)
        if "SSS" in raw_fmt: raw_fmt = raw_fmt.replace("SSS", ms_str)
        elif "SS" in raw_fmt: raw_fmt = raw_fmt.replace("SS", ms_str[:2])
        elif "S" in raw_fmt: raw_fmt = raw_fmt.replace("S", ms_str[:1])
        py_fmt = self.excel_to_py_format(raw_fmt)
        if "%H" in py_fmt and not self.is_24h:
            h = now.hour % 12
            py_fmt = py_fmt.replace("%H", str(12 if h == 0 else h))
        main_time = now.strftime(py_fmt)
        if not self.is_24h and self.show_ampm and "hh" in self.excel_date_format:
            main_time += " " + now.strftime("%p")
        if self.show_weekday:
            weeks = ["週一","週二","週三","週四","週五","週六","週日"] if self.weekday_lang=="ZH" else ["Mon","Tue","Wed","Thu","Fri","Sat","Sun"]
            main_time += f" ({weeks[now.weekday()]})"
        if self.show_timezone:
            main_time += f" {self.format_utc_offset()}"
        ts = now.timestamp()
        self.current_ts_val, self.current_sec_val = str(int(ts*1000)), str(int(ts))
        display_text = main_time
        if self.show_timestamp:
            ts_str = self.current_ts_val if self.show_ms else self.current_sec_val
            display_text += f"\n{ts_str}"
        self.label.config(text=display_text)
        has_ms = any(x in self.excel_date_format for x in ["S", "SSS", "SS"])
        self.root.after(50 if (has_ms or (self.show_timestamp and self.show_ms)) else 1000, self.update_clock)

    # --- 100% 還原選單結構 ---
    def open_ctk_main_menu(self, event):
        if hasattr(self, "_ctk_main_menu_win") and self._ctk_main_menu_win and self._ctk_main_menu_win.winfo_exists():
            self._ctk_main_menu_win.destroy()
        if hasattr(self, "_ctk_submenu_win") and self._ctk_submenu_win and self._ctk_submenu_win.winfo_exists():
            self._ctk_submenu_win.destroy()
        self._build_ctk_menu(event)

    def _build_ctk_menu(self, event):
        try:
            win = ctk.CTkToplevel(self.root)
            # overrideredirect 選單不需要工作列圖標，略過 set_win_icon 加速開啟
            self._ctk_main_menu_win = win
            win.overrideredirect(True)
            win.attributes("-topmost", True)
            win.withdraw()  # 先隱藏，避免定位前閃現


            frame = ctk.CTkFrame(win, corner_radius=14, border_width=1)
            frame.pack(padx=3, pady=3)
            btn_f = ("Microsoft JhengHei", 12)
            btn_h = 24
            top_label = self.t("top_on") if self.is_topmost else self.t("top_off")
            labels = [
                self.t("copy_display"),
                self.t("copy_ts"),
                top_label,
                f"{'格式設定' if self.language=='ZH' else 'Format'}   ▸",
                f"{'樣式設定' if self.language=='ZH' else 'Style'}   ▸",
                f"{'設定' if self.language=='ZH' else 'Settings'}   ▸",
                self.t("ts_converter"),
                self.t("reset_style"),
                self.t("exit"),
            ]
            btn_w = self.calc_ctk_menu_width(labels, font_tuple=btn_f, extra_px=32, min_px=140, max_px=260)
            c = self.get_menu_colors()
            # 用 tk.Frame 撐寬度（比 CTkFrame 更可靠）
            try:
                raw_fbg = frame.cget("fg_color")
                fbg = raw_fbg[1] if self.appearance_mode == "dark" else raw_fbg[0]
            except:
                fbg = "#2B2B2B" if self.appearance_mode == "dark" else "#EBEBEB"
            _sizer = tk.Frame(frame, width=btn_w, height=1, bg=fbg)
            _sizer.pack()
            _sizer.pack_propagate(False)
            # 深色分隔線 #4A4A4A，淽色分隔線 #BBBBBB
            sep_color = "#4A4A4A" if self.appearance_mode == "dark" else "#BBBBBB"
            danger_text_c = "#B71C1C" if self.appearance_mode == "light" else c["fg"]

            def add_sep():
                tk.Frame(frame, bg=sep_color, height=1).pack(fill="x", padx=12, pady=3)

            def safe_destroy(w):
                try:
                    if w and w.winfo_exists():
                        w.destroy()
                except:
                    pass

            # 用 dict 包裝讓後面的重定義能被所有 lambda 捕捉到
            _cm = {"fn": None}

            def close_menu():
                if _cm["fn"]: _cm["fn"]()

            def _base_close():
                safe_destroy(win)
                self._menu_poll_mute_ref = None
                for attr in ("_ctk_submenu_win", "_ctk_submenu2_win"):
                    if hasattr(self, attr):
                        safe_destroy(getattr(self, attr))

            _cm["fn"] = _base_close

            def run_action(action):
                close_menu()
                self.root.after(20, action)


            def add_btn(text, cmd, *, danger=False, pady=(0, 0), icon="", icon_color=None, right_icon=""):
                hover_c = "#FFCDD2" if (danger and self.appearance_mode == "light") else ("#C0392B" if danger else c["active_bg"])
                danger_fg = "#B71C1C" if self.appearance_mode == "light" else "#FF6B6B"
                text_c = danger_fg if danger else c["fg"]
                ic = icon_color or text_c

                outer = ctk.CTkFrame(frame, fg_color="transparent", corner_radius=8, cursor="hand2")
                outer.pack(fill="x", padx=4, pady=pady)
                outer.columnconfigure(1, weight=1)

                # 欄0：icon
                lbl_ic = ctk.CTkLabel(outer, text=icon, font=("Segoe UI Emoji", 11),
                                      text_color=ic, fg_color="transparent", width=26, anchor="center")
                lbl_ic.grid(row=0, column=0, padx=(4, 0), pady=2)

                # 欄1：文字
                lbl_txt = ctk.CTkLabel(outer, text=text, font=btn_f,
                                       text_color=text_c, fg_color="transparent", anchor="w")
                lbl_txt.grid(row=0, column=1, sticky="ew", padx=(2, 0), pady=2)

                # 欄2：右側 icon
                lbl_ri = ctk.CTkLabel(outer, text=right_icon, font=btn_f,
                                      text_color=text_c, fg_color="transparent", width=24, anchor="center")
                lbl_ri.grid(row=0, column=2, padx=(0, 4))

                def on_enter(e):
                    outer.configure(fg_color=hover_c)
                def on_leave(e):
                    outer.configure(fg_color="transparent")
                def on_click(e):
                    if cmd: cmd()
                    return "break"

                for w in (outer, lbl_ic, lbl_txt, lbl_ri):
                    w.bind("<Enter>", on_enter)
                    w.bind("<Leave>", on_leave)
                    w.bind("<Button-1>", on_click)

                return outer

            _sub_lock = [False]

            def open_fmt_sub(btn):
                if _sub_lock[0]: return
                _sub_lock[0] = True
                _poll_mute[0] = 10
                win.after(80, lambda: _sub_lock.__setitem__(0, False))
                mm = self._mainmenu_rect
                bx = mm[2] if mm else win.winfo_rootx() + win.winfo_width()
                by = btn.winfo_rooty()
                self.show_format_submenu(bx, by, close_menu)

            def open_style_sub(btn):
                if _sub_lock[0]: return
                _sub_lock[0] = True
                _poll_mute[0] = 10
                win.after(80, lambda: _sub_lock.__setitem__(0, False))
                mm = self._mainmenu_rect
                bx = mm[2] if mm else win.winfo_rootx() + win.winfo_width()
                by = btn.winfo_rooty()
                self.show_style_submenu(bx, by, close_menu)

            def open_settings_sub(btn):
                if _sub_lock[0]: return
                _sub_lock[0] = True
                _poll_mute[0] = 10
                win.after(80, lambda: _sub_lock.__setitem__(0, False))
                mm = self._mainmenu_rect
                bx = mm[2] if mm else win.winfo_rootx() + win.winfo_width()
                by = btn.winfo_rooty()
                self.show_settings_submenu(bx, by, close_menu)

            is_zh = self.language == "ZH"
            fmt_label   = "格式設定" if is_zh else "Format"
            style_label = "樣式設定" if is_zh else "Style"
            settings_label = "設定" if is_zh else "Settings"

            # 置頂 — 圖釘 icon
            pin_icon = "📌" if self.is_topmost else "📍"
            add_btn(top_label, lambda: [self.toggle_topmost(), close_menu()],
                    pady=(4, 0), icon=pin_icon, icon_color="#E67E22")
            add_sep()
            # 複製 — 無 icon
            add_btn(self.t("copy_display"), lambda: [self.root.clipboard_clear(), self.root.clipboard_append(self.label.cget("text")), close_menu()])
            add_btn(self.t("copy_ts"), lambda: [self.copy_smart_timestamp(), close_menu()])
            add_sep()
            # 格式設定 — 🗓 icon，▸ 靠右
            fmt_btn = add_btn(fmt_label, lambda: open_fmt_sub(fmt_btn_ref), icon="🗓", right_icon="▸")
            fmt_btn_ref = fmt_btn

            style_btn = add_btn(style_label, lambda: open_style_sub(style_btn_ref), icon="🖌", right_icon="▸")
            style_btn_ref = style_btn

            add_sep()
            add_btn(self.t("ts_converter"), lambda: run_action(self.open_converter))
            add_sep()
            settings_btn = add_btn(settings_label, lambda: open_settings_sub(settings_btn_ref), icon="🔧", icon_color="#85C1E9", right_icon="▸")
            settings_btn_ref = settings_btn

            add_btn(self.t("reset_style"), lambda: [self.apply_default(), close_menu()])
            add_btn(self.t("exit"), lambda: [close_menu(), self.safe_exit()],
                    danger=True, pady=(0, 4), right_icon="⏻")

            win.bind("<Escape>", lambda e: close_menu())

            # 計算尺寸：withdraw 狀態下先 update_idletasks 取正確 reqheight
            win.update_idletasks()
            menu_w = win.winfo_reqwidth()
            menu_h = win.winfo_reqheight()
            vsx, vsy, vsw, vsh = self._get_virtual_screen()

            mx = event.x_root
            my = event.y_root
            if mx + menu_w > vsx + vsw:
                mx = vsx + vsw - menu_w
            if my + menu_h > vsy + vsh:
                my = max(vsy, my - menu_h)
            mx = max(vsx, mx)
            my = max(vsy, my)
            win.geometry(f"+{mx}+{my}")
            win.deiconify()  # 定位完成後才顯示
            # cache 主選單 rect，供 poll hit-test 使用（避免 poll 呼叫 winfo）
            self._mainmenu_rect = (mx, my, mx + menu_w, my + menu_h)
            self._submenu_rect = None  # 清除舊的 submenu rect

            def _cleanup_poll():
                if hasattr(self, "_menu_poll_id") and self._menu_poll_id:
                    try: self.root.after_cancel(self._menu_poll_id)
                    except: pass
                    self._menu_poll_id = None

            _cm["fn"] = lambda: [_cleanup_poll(), _base_close()]

            # 輪詢滑鼠按鍵，若點到 win/submenu 外就關閉（支援任意視窗包括其他程式）
            _poll_mute = [0]  # 靜默計數器，開啟次選單後暫停幾次 poll
            self._menu_poll_mute_ref = _poll_mute  # 讓次選單可以觸發靜默

            VK_LBUTTON = 0x01
            VK_RBUTTON = 0x02
            _btn_was_up = [True, True]

            def _poll_click():
                # 用 cache rect 做 hit-test，避免每 tick 呼叫 winfo_rootx 等 Tk IPC
                if _poll_mute[0] > 0:
                    _poll_mute[0] -= 1
                    self._menu_poll_id = self.root.after(60, _poll_click)
                    return
                try:
                    lbtn = bool(ctypes.windll.user32.GetAsyncKeyState(VK_LBUTTON) & 0x8000)
                    rbtn = bool(ctypes.windll.user32.GetAsyncKeyState(VK_RBUTTON) & 0x8000)
                    states = [lbtn, rbtn]
                    any_new_click = False
                    for i, pressed in enumerate(states):
                        if pressed and _btn_was_up[i]:
                            any_new_click = True
                        _btn_was_up[i] = not pressed

                    if any_new_click:
                        pt = ctypes.wintypes.POINT()
                        ctypes.windll.user32.GetCursorPos(ctypes.byref(pt))
                        cx, cy = pt.x, pt.y
                        inside = False
                        for rect in (self._mainmenu_rect, self._submenu_rect):
                            if rect and rect[0] <= cx <= rect[2] and rect[1] <= cy <= rect[3]:
                                inside = True
                                break
                        if not inside:
                            close_menu()
                            return
                except:
                    pass
                self._menu_poll_id = self.root.after(60, _poll_click)

            self._menu_poll_id = self.root.after(150, _poll_click)

            win.bind("<Escape>", lambda e: close_menu())
        except:
            pass

    def _get_virtual_screen(self):
        """取得虛擬螢幕範圍，結果快取在 instance 避免重複呼叫 GetSystemMetrics。"""
        if not hasattr(self, "_vscreen_cache"):
            self._vscreen_cache = None
        if self._vscreen_cache is None:
            try:
                vsx = ctypes.windll.user32.GetSystemMetrics(76)
                vsy = ctypes.windll.user32.GetSystemMetrics(77)
                vsw = ctypes.windll.user32.GetSystemMetrics(78)
                vsh = ctypes.windll.user32.GetSystemMetrics(79)
                self._vscreen_cache = (vsx, vsy, vsw, vsh)
            except:
                w = self.root.winfo_screenwidth()
                h = self.root.winfo_screenheight()
                self._vscreen_cache = (0, 0, w, h)
        return self._vscreen_cache

    def _position_submenu(self, sub_win, x, y):
        """
        次選單定位：優先往主選單右側緊黏展開，右側不夠時往左緊黏展開。
        x = 主選單右邊緣（已由呼叫者從 _mainmenu_rect[2] 取得）
        y = 按鈕的 rooty
        """
        # 必須先 update_idletasks，否則 withdraw 狀態的 reqwidth/reqheight 可能回傳 1
        sub_win.update_idletasks()

        vsx, vsy, vsw, vsh = self._get_virtual_screen()
        vsr = vsx + vsw
        vsb = vsy + vsh

        mw = sub_win.winfo_reqwidth()
        mh = sub_win.winfo_reqheight()

        mm = getattr(self, "_mainmenu_rect", None)
        main_left  = mm[0] if mm else (x - 200)
        main_right = mm[2] if mm else x

        # --- 水平方向：緊黏，無間距 ---
        if main_right + mw <= vsr:
            fx = main_right          # 右側展開，緊黏主選單右邊緣
        else:
            fx = main_left - mw      # 左側展開，緊黏主選單左邊緣
        fx = max(vsx, min(fx, vsr - mw))

        # --- 垂直方向：從按鈕位置往下；超出螢幕時往上貼齊 ---
        if y + mh <= vsb:
            fy = y
        else:
            fy = vsb - mh
        fy = max(vsy, fy)

        sub_win.geometry(f"+{fx}+{fy}")
        sub_win.deiconify()
        self._submenu_rect = (fx, fy, fx + mw, fy + mh)

    def show_format_submenu(self, x, y, close_main=None):
        if hasattr(self, "_ctk_submenu_win") and self._ctk_submenu_win and self._ctk_submenu_win.winfo_exists():
            self._ctk_submenu_win.destroy()

        win = ctk.CTkToplevel(self.root)
        # overrideredirect 選單不需要圖標
        self._ctk_submenu_win = win
        win.withdraw()
        win.overrideredirect(True)
        win.attributes("-topmost", True)

        frame = ctk.CTkFrame(win, corner_radius=14, border_width=1)
        frame.pack(padx=3, pady=3)
        btn_f = ("Microsoft JhengHei", 12)
        btn_h = 24
        sys_offset = self.get_system_utc_offset()
        labels = [
            self.t('fmt_12h') if self.is_24h else self.t('fmt_24h'),
            self.t('hide_ampm') if self.show_ampm else self.t('show_ampm'),
            self.t('hide_weekday') if self.show_weekday else self.t('show_weekday'),
            self.t('week_en') if self.weekday_lang=="ZH" else self.t('week_zh'),
            self.t('hide_tz') if self.show_timezone else self.t('show_tz'),
            self.t('change_tz'),
            f"{self.t('reset_sys_tz')} ({self.format_utc_offset(sys_offset)})",
            self.t('custom_fmt'),
            self.t('hide_ts') if self.show_timestamp else self.t('show_ts'),
            self.t('show_sec') if self.show_ms else self.t('show_ms'),
        ]
        btn_w = self.calc_ctk_menu_width(labels, font_tuple=btn_f, extra_px=24, min_px=140, max_px=280)
        try:
            raw_fbg = frame.cget("fg_color")
            fbg = raw_fbg[1] if self.appearance_mode == "dark" else raw_fbg[0]
        except:
            fbg = "#2B2B2B" if self.appearance_mode == "dark" else "#EBEBEB"
        _sizer = tk.Frame(frame, width=btn_w, height=1, bg=fbg)
        _sizer.pack(); _sizer.pack_propagate(False)
        c = self.get_menu_colors()
        sep_color = "#4A4A4A" if self.appearance_mode == "dark" else "#BBBBBB"

        def safe_close():
            try:
                if win.winfo_exists():
                    win.destroy()
            except:
                pass
            self._submenu_rect = None  # 清除 cache，poll 不再偵測已關閉的 submenu
            if close_main:
                close_main()

        def run_action(action):
            safe_close()
            self.root.after(20, action)

        def _build_buttons():
            # 清除 frame 內現有按鈕（保留 _sizer）
            for w in list(frame.winfo_children()):
                if w is not _sizer:
                    try: w.destroy()
                    except: pass

            def toggle_and_refresh(attr, val_fn=None):
                if val_fn:
                    setattr(self, attr, val_fn())
                else:
                    setattr(self, attr, not getattr(self, attr))
                if hasattr(self, '_menu_poll_mute_ref') and self._menu_poll_mute_ref:
                    try: self._menu_poll_mute_ref[0] = 8
                    except: pass
                _build_buttons()

            def add_btn(text, cmd, *, pady=(0, 0), icon=""):
                hover_c = c["active_bg"]
                outer = ctk.CTkFrame(frame, fg_color="transparent", corner_radius=8, cursor="hand2")
                outer.pack(fill="x", padx=4, pady=pady)
                outer.columnconfigure(1, weight=1)
                lbl_ic = ctk.CTkLabel(outer, text=icon, font=("Segoe UI Emoji", 11),
                                   text_color=c["fg"], fg_color="transparent", width=26, anchor="center")
                lbl_ic.grid(row=0, column=0, padx=(4, 0), pady=2)
                lbl_txt = ctk.CTkLabel(outer, text=text, font=btn_f,
                                       text_color=c["fg"], fg_color="transparent", anchor="w")
                lbl_txt.grid(row=0, column=1, sticky="ew", padx=(2, 8), pady=2)
                def on_enter(e): outer.configure(fg_color=hover_c)
                def on_leave(e): outer.configure(fg_color="transparent")
                def on_click(e): cmd(); return "break"
                for w in (outer, lbl_ic, lbl_txt):
                    w.bind("<Enter>", on_enter)
                    w.bind("<Leave>", on_leave)
                    w.bind("<Button-1>", on_click)
                return outer

            def add_sep():
                tk.Frame(frame, bg=sep_color, height=1).pack(fill="x", padx=12, pady=3)

            # 群組1：時間格式 + 星期 + 時間戳
            add_btn(self.t('fmt_12h') if self.is_24h else self.t('fmt_24h'),
                    lambda: toggle_and_refresh('is_24h'), pady=(4, 0), icon="🕐")
            if not self.is_24h:
                add_btn(self.t('hide_ampm') if self.show_ampm else self.t('show_ampm'),
                        lambda: toggle_and_refresh('show_ampm'))
            add_btn(self.t('hide_weekday') if self.show_weekday else self.t('show_weekday'),
                    lambda: toggle_and_refresh('show_weekday'), icon="📅")
            if self.show_weekday:
                add_btn(self.t('week_en') if self.weekday_lang=="ZH" else self.t('week_zh'),
                        lambda: toggle_and_refresh('weekday_lang', lambda: "EN" if self.weekday_lang=="ZH" else "ZH"))
            add_btn(self.t('hide_ts') if self.show_timestamp else self.t('show_ts'),
                    lambda: toggle_and_refresh('show_timestamp'), icon="⏱")
            if self.show_timestamp:
                add_btn(self.t('show_sec') if self.show_ms else self.t('show_ms'),
                        lambda: toggle_and_refresh('show_ms'))
            add_sep()
            # 群組2：時區
            add_btn(self.t('hide_tz') if self.show_timezone else self.t('show_tz'),
                    lambda: toggle_and_refresh('show_timezone'), icon="🌍")
            add_btn(self.t('change_tz'), lambda: run_action(self.change_utc_dialog))
            add_btn(f"{self.t('reset_sys_tz')} ({self.format_utc_offset(sys_offset)})",
                    lambda: toggle_and_refresh('utc_offset', lambda: sys_offset))
            add_sep()
            # 群組3：自定義格式
            add_btn(self.t('custom_fmt'), lambda: run_action(self.change_format_dialog), pady=(0, 4), icon="✏")

        _build_buttons()

        win.bind("<Escape>", lambda e: safe_close())
        self._position_submenu(win, x, y)

    def show_style_submenu(self, x, y, close_main=None):
        if hasattr(self, "_ctk_submenu_win") and self._ctk_submenu_win and self._ctk_submenu_win.winfo_exists():
            self._ctk_submenu_win.destroy()
        win = ctk.CTkToplevel(self.root)
        # overrideredirect 選單不需要圖標
        self._ctk_submenu_win = win
        win.withdraw()
        win.overrideredirect(True)
        win.attributes("-topmost", True)

        frame = ctk.CTkFrame(win, corner_radius=14, border_width=1)
        frame.pack(padx=3, pady=3)
        btn_f = ("Microsoft JhengHei", 12)
        btn_h = 24
        labels = [self.t("change_fg"), self.t("change_bg")]
        btn_w = self.calc_ctk_menu_width(labels, font_tuple=btn_f, extra_px=24, min_px=130, max_px=240)
        try:
            raw_fbg = frame.cget("fg_color")
            fbg = raw_fbg[1] if self.appearance_mode == "dark" else raw_fbg[0]
        except:
            fbg = "#2B2B2B" if self.appearance_mode == "dark" else "#EBEBEB"
        _sizer = tk.Frame(frame, width=btn_w, height=1, bg=fbg)
        _sizer.pack(); _sizer.pack_propagate(False)
        c = self.get_menu_colors()

        def safe_close():
            try:
                if win.winfo_exists(): win.destroy()
            except: pass
            self._submenu_rect = None
            if close_main: close_main()

        def run_action(action):
            safe_close()
            self.root.after(20, action)

        def add_btn(text, cmd, *, pady=(0, 0)):
            b = ctk.CTkButton(frame, text=text, height=btn_h, font=btn_f, anchor="w",
                fg_color="transparent", hover_color=c["active_bg"],
                text_color=c["fg"], corner_radius=8, command=cmd)
            b.pack(padx=4, pady=pady, fill="x")
            return b

        add_btn(self.t("change_fg"), lambda: run_action(self.change_text_color), pady=(4, 0))
        add_btn(self.t("change_bg"), lambda: run_action(self.change_bg_color), pady=(0, 4))
        win.bind("<Escape>", lambda e: safe_close())
        self._position_submenu(win, x, y)

    def show_settings_submenu(self, x, y, close_main=None):
        if hasattr(self, "_ctk_submenu_win") and self._ctk_submenu_win and self._ctk_submenu_win.winfo_exists():
            self._ctk_submenu_win.destroy()
        win = ctk.CTkToplevel(self.root)
        # overrideredirect 選單不需要圖標
        self._ctk_submenu_win = win
        win.withdraw()
        win.overrideredirect(True)
        win.attributes("-topmost", True)

        frame = ctk.CTkFrame(win, corner_radius=14, border_width=1)
        frame.pack(padx=3, pady=3)
        btn_f = ("Microsoft JhengHei", 12)
        btn_h = 24
        lang_item = self.t("lang_to_en") if self.language == "ZH" else self.t("lang_to_zh")
        theme_toggle = self.t("use_light_theme") if self.appearance_mode == "dark" else self.t("use_dark_theme")
        labels = [lang_item, theme_toggle]
        btn_w = self.calc_ctk_menu_width(labels, font_tuple=btn_f, extra_px=24, min_px=130, max_px=240)
        try:
            raw_fbg = frame.cget("fg_color")
            fbg = raw_fbg[1] if self.appearance_mode == "dark" else raw_fbg[0]
        except:
            fbg = "#2B2B2B" if self.appearance_mode == "dark" else "#EBEBEB"
        _sizer = tk.Frame(frame, width=btn_w, height=1, bg=fbg)
        _sizer.pack(); _sizer.pack_propagate(False)
        c = self.get_menu_colors()

        def safe_close():
            try:
                if win.winfo_exists(): win.destroy()
            except: pass
            self._submenu_rect = None
            if close_main: close_main()


        def add_btn(text, cmd, *, pady=(0, 0), icon="", icon_color=None):
            hover_c = c["active_bg"]
            ic = icon_color or c["fg"]
            outer = ctk.CTkFrame(frame, fg_color="transparent", corner_radius=8, cursor="hand2")
            outer.pack(fill="x", padx=4, pady=pady)
            outer.columnconfigure(1, weight=1)

            lbl_ic = ctk.CTkLabel(outer, text=icon, font=("Segoe UI Emoji", 11),
                                  text_color=ic, fg_color="transparent", width=26, anchor="center")
            lbl_ic.grid(row=0, column=0, padx=(4, 0), pady=2)

            lbl_txt = ctk.CTkLabel(outer, text=text, font=btn_f,
                                   text_color=c["fg"], fg_color="transparent", anchor="w")
            lbl_txt.grid(row=0, column=1, sticky="ew", padx=(2, 8), pady=2)

            def on_enter(e):
                outer.configure(fg_color=hover_c)
            def on_leave(e):
                outer.configure(fg_color="transparent")
            def on_click(e):
                cmd(); return "break"

            for w in (outer, lbl_ic, lbl_txt):
                w.bind("<Enter>", on_enter)
                w.bind("<Leave>", on_leave)
                w.bind("<Button-1>", on_click)
            return outer

        VERSION = "Ver 1.0.1"
        # 當前深色→按鈕說「使用淺色」→顯示🌤；當前淺色→按鈕說「使用深色」→顯示🌙
        # 當前深色→「使用淺色」→顯示☀️；當前淺色→「使用深色」→顯示🌙
        moon_icon = "☀️" if self.appearance_mode == "dark" else "🌙"
        def do_lang_toggle():
            safe_close()
            if close_main: close_main()
            self.root.after(30, self.toggle_language)

        add_btn(lang_item, do_lang_toggle, pady=(4, 0), icon="🌐")

        def do_theme_toggle():
            # 先關閉所有選單，再切換主題，避免 CTk 全域重繪時選單閃爍
            safe_close()
            if close_main: close_main()
            self.root.after(30, self.toggle_theme_mode)

        add_btn(theme_toggle, do_theme_toggle,
                icon=moon_icon, icon_color="#F1C40F" if self.appearance_mode == "dark" else "#1A3A6B")

        def open_help():
            safe_close()
            self.root.after(20, self._open_help_dialog)

        add_btn(self.t("help") if hasattr(self, "t") else "說明", open_help, pady=(0, 4), icon="❓", icon_color="#E74C3C")
        win.bind("<Escape>", lambda e: safe_close())
        self._position_submenu(win, x, y)

    def _open_help_dialog(self):
        VERSION = "Ver 1.0.1"
        BUILD = "Dev-0406.3"  # Dev 版號（供開發紀錄用）
        win = ctk.CTkToplevel(self.root)
        self.set_win_icon(win)
        win.title(self.t("help") + " / About")
        win.attributes("-topmost", True)
        win.resizable(False, False)
        win.geometry(self.get_popup_pos(300))
        win_f = ("Microsoft JhengHei", 13)
        pad = ctk.CTkFrame(win, fg_color="transparent")
        pad.pack(padx=24, pady=20, fill="both", expand=True)

        # 版本號
        ctk.CTkLabel(pad, text=VERSION, font=("Microsoft JhengHei", 15, "bold"),
                     anchor="center").pack(fill="x", pady=(0, 14))

        def section(title, lines):
            ctk.CTkLabel(pad, text=title, font=("Microsoft JhengHei", 13, "bold"),
                         anchor="w").pack(fill="x", pady=(8, 2))
            for line in lines:
                ctk.CTkLabel(pad, text=line, font=win_f, anchor="w",
                             justify="left").pack(fill="x", padx=8)

        if self.language == "ZH":
            section("🖱 滑鼠操作", [
                "滾輪上/下          調整整體透明度 (±5%)",
                "Ctrl + 滾輪上     放大字體",
                "Ctrl + 滾輪下     縮小字體",
                "左鍵拖曳           移動時鐘",
                "左鍵雙擊           關閉時鐘",
                "右鍵點擊           開啟選單",
            ])
        else:
            section("🖱 Mouse", [
                "Scroll up/down         Adjust opacity (±5%)",
                "Ctrl + Scroll up       Increase font size",
                "Ctrl + Scroll down     Decrease font size",
                "Left drag              Move clock",
                "Left double-click      Close clock",
                "Right click            Open menu",
            ])

        ctk.CTkButton(pad, text=self.t("msg_ok"), font=win_f, width=90,
                      command=win.destroy).pack(pady=(18, 0))

        # Build 號顯示在視窗最下方
        hint_color = "#666666" if self.appearance_mode == "dark" else "#999999"
        ctk.CTkLabel(pad, text=BUILD, font=("Consolas", 10),
                     text_color=hint_color, anchor="center").pack(pady=(10, 0))

    def _set_app_icon(self):
        """載入 app.ico，設定主視窗工作列圖標"""
        import sys, os
        # 依序嘗試各種路徑（PyInstaller / Nuitka / 直接執行）
        candidates = []
        if hasattr(sys, "_MEIPASS"):                          # PyInstaller
            candidates.append(sys._MEIPASS)
        if hasattr(sys, "frozen"):                            # Nuitka onefile
            candidates.append(os.path.dirname(sys.executable))
        candidates.append(os.path.dirname(os.path.abspath(__file__)))
        candidates.append(os.path.dirname(os.path.abspath(sys.argv[0])))

        ico_path = None
        for base in candidates:
            p = os.path.join(base, "app.ico")
            if os.path.exists(p):
                ico_path = p
                break

        if not ico_path:
            self._app_icon = None
            return

        self._app_icon = ico_path
        try:
            self.root.wm_iconbitmap(ico_path)
        except: pass
        # overrideredirect 視窗工作列圖標需要額外設定
        try:
            self.root.iconbitmap(ico_path)
        except: pass
        # 延遲再設一次確保生效
        self.root.after(100, lambda: self._apply_taskbar_icon(ico_path))

    def _apply_taskbar_icon(self, ico_path):
        """確保工作列圖標正確顯示（overrideredirect 視窗需特殊處理）"""
        try:
            self.root.wm_iconbitmap(ico_path)
            self.root.iconbitmap(ico_path)
        except: pass

    def set_win_icon(self, win):
        """為任意 CTkToplevel 設定圖標"""
        if not self._app_icon:
            return
        ico = self._app_icon
        try:
            win.after(1,   lambda: win.wm_iconbitmap(ico) if win.winfo_exists() else None)
            win.after(200, lambda: win.wm_iconbitmap(ico) if win.winfo_exists() else None)
            win.after(500, lambda: win.wm_iconbitmap(ico) if win.winfo_exists() else None)
        except: pass

    def toggle_topmost(self):
        self.is_topmost = not self.is_topmost
        self.root.attributes("-topmost", self.is_topmost)

    # 語言切換 ZH <-> EN
    def toggle_language(self):
        self.language = "EN" if getattr(self, "language", "ZH") == "ZH" else "ZH"

    def toggle_theme_mode(self):
        self.appearance_mode = "light" if self.appearance_mode == "dark" else "dark"
        self.apply_appearance_mode()

    def on_font_zoom(self, e):
        self.font_size = max(10, min(100, self.font_size + (2 if e.delta > 0 else -2)))
        self.refresh_clock_font()

    # --- 透明度 (修正保護機制與同步) ---
    def open_converter(self):
        conv = ctk.CTkToplevel(self.root)
        self.set_win_icon(conv)
        conv.title(self.t("converter_title"))
        conv.attributes("-topmost", True)
        conv.geometry(self.get_popup_pos(200))
        win_f = ('Microsoft JhengHei', 14)

        main_container = ctk.CTkFrame(conv, fg_color="transparent")
        main_container.pack(padx=20, pady=(10, 10))

        # 時區輸入設定
        tz_offsets = list(range(-12, 15))
        def fmt_tz(h): return f"+{h}" if h >= 0 else str(h)
        init_h = int(round(self.utc_offset))
        init_h = max(-12, min(14, init_h))
        _conv_tz = [init_h]
        tz_entry_var = tk.StringVar(value=fmt_tz(init_h))

        def clamp_and_apply(val_str):
            s = val_str.strip()
            if s in ("", "+", "-"): return
            try:
                h = int(s)
                h = max(-12, min(14, h))
                _conv_tz[0] = h
                tz_entry_var.set(fmt_tz(h))
            except: pass

        def on_tz_validate(P):
            if P == "" or P in ("+", "-"): return True
            try: int(P); return True
            except: return False

        hint_color = "#888888"
        hint_f = ("Microsoft JhengHei", 11)

        # 標題列
        ctk.CTkLabel(main_container, text="日期時間" if self.language=="ZH" else "Datetime",
                     font=win_f, anchor="w").grid(row=0, column=0, padx=(10,0), pady=(10,0), sticky="w")
        ctk.CTkLabel(main_container, text="UTC",
                     font=win_f, anchor="center").grid(row=0, column=1, padx=(8,8), pady=(10,0))
        ctk.CTkLabel(main_container, text=self.t("converter_label_ts"),
                     font=win_f, anchor="w").grid(row=0, column=2, padx=(20,10), pady=(10,0), sticky="w")

        # 輸入列
        e_date_ctk = ctk.CTkEntry(main_container, font=win_f, width=185)
        e_date_ctk.grid(row=1, column=0, padx=(10,0), pady=(4,0), sticky="w")
        e_date = e_date_ctk._entry

        vcmd = (conv.register(on_tz_validate), "%P")
        tz_entry = ctk.CTkEntry(main_container, textvariable=tz_entry_var,
                                width=45, font=win_f, justify="center",
                                validate="key", validatecommand=vcmd)
        tz_entry.grid(row=1, column=1, padx=(8,8), pady=(4,0))

        e_ts_ctk = ctk.CTkEntry(main_container, font=win_f, width=125)
        e_ts_ctk.grid(row=1, column=2, padx=(20,10), pady=(4,0), sticky="w")
        e_ts = e_ts_ctk._entry

        # 提示文字列（只顯示日期格式）
        ctk.CTkLabel(main_container, text="YYYY-MM-DD hh:mm:ss.SSS",
                     font=hint_f, text_color=hint_color, anchor="w").grid(
                     row=2, column=0, padx=(10,0), pady=(2,10), sticky="w")

        tz_entry._entry.bind("<FocusOut>", lambda e: clamp_and_apply(tz_entry_var.get()))
        tz_entry._entry.bind("<Return>", lambda e: clamp_and_apply(tz_entry_var.get()))

        def on_tz_wheel(e):
            try: cur = int(tz_entry_var.get() or "0")
            except: cur = _conv_tz[0]
            new_h = max(-12, min(14, cur + (-1 if e.delta > 0 else 1)))
            _conv_tz[0] = new_h
            tz_entry_var.set(fmt_tz(new_h))

        tz_entry.bind("<MouseWheel>", on_tz_wheel)
        tz_entry._entry.bind("<MouseWheel>", on_tz_wheel)

        def fetch_now():
            clamp_and_apply(tz_entry_var.get())
            tz = timezone(timedelta(hours=_conv_tz[0]))
            now = datetime.now(tz)
            e_date.delete(0, tk.END); e_date.insert(0, now.strftime("%Y-%m-%d %H:%M:%S.%f")[:-3])
            e_ts.delete(0, tk.END); e_ts.insert(0, str(int(now.timestamp()*1000)))

        fetch_now()
        self.last_f = e_date
        e_date.bind("<FocusIn>", lambda e: setattr(self, 'last_f', e_date))
        e_ts.bind("<FocusIn>", lambda e: setattr(self, 'last_f', e_ts))

        def do_conv():
            try:
                clamp_and_apply(tz_entry_var.get())
                tz = timezone(timedelta(hours=_conv_tz[0]))

                # 來源：時間戳 -> 目的：日期時間
                focus_w = self.root.focus_get()
                if focus_w == e_ts:
                    self.last_f = e_ts
                elif focus_w == e_date:
                    self.last_f = e_date

                if self.last_f == e_ts:
                    val = "".join(re.findall(r"[0-9]", e_ts.get().strip()))
                    if not val:
                        return

                    # 嚴格依碼長判定：10 = 秒, 13 = 毫秒
                    if len(val) == 10:
                        seconds = int(val)
                        dt = datetime.fromtimestamp(seconds, tz)
                        out = dt.strftime("%Y-%m-%d %H:%M:%S")
                    elif len(val) == 13:
                        ms = int(val)
                        seconds = ms // 1000
                        microsecond = (ms % 1000) * 1000  # 精準還原毫秒
                        dt = datetime.fromtimestamp(seconds, tz).replace(microsecond=microsecond)
                        out = dt.strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]  # 只取 SSS(3 碼)
                    else:
                        self.show_ctk_message(self.t("msg_hint"), self.t("err_ts_len"))
                        return

                    e_date.delete(0, tk.END)
                    e_date.insert(0, out)

                # 來源：日期時間 -> 目的：時間戳
                else:
                    s = e_date.get().strip()
                    if not s:
                        return

                    # 日期分隔符允許 / 或 ,
                    c = re.sub(r"[/,]", "-", s)

                    # 只要「小數點後有數字」就視為有毫秒
                    has_ms = bool(re.search(r"\.\d+$", s))

                    dt = datetime.strptime(
                        c + (".000" if not has_ms else ""),
                        "%Y-%m-%d %H:%M:%S.%f",
                    ).replace(tzinfo=tz)

                    sec = int(dt.timestamp())
                    e_ts.delete(0, tk.END)
                    if has_ms:
                        # 用整數組合避免 float 造成的捨入誤差
                        e_ts.insert(0, str(sec * 1000 + dt.microsecond // 1000))
                    else:
                        e_ts.insert(0, str(sec))
            except: 
                self.show_ctk_message(self.t("msg_hint"), self.t("err_format"))

        btn_frame = ctk.CTkFrame(conv, fg_color="transparent")
        btn_frame.pack(pady=(8, 14))
        btn_fetch = ctk.CTkButton(btn_frame, text=self.t("btn_fetch_now"), font=win_f, command=fetch_now)
        btn_fetch.pack(side="left", padx=8)
        btn_conv = ctk.CTkButton(btn_frame, text=self.t("btn_start_conv"), font=win_f,
                                 fg_color="#2FA572", hover_color="#106A43", command=do_conv)
        btn_conv.pack(side="left", padx=8)

        conv.resizable(False, False)

    def change_format_dialog(self):
        dialog = ctk.CTkToplevel(self.root)
        self.set_win_icon(dialog)
        dialog.title(self.t("fmt_dialog_title"))
        dialog.attributes("-topmost", True)
        dialog.geometry(self.get_popup_pos(200))
        dialog.resizable(False, False)
        # 英文提示較長，視窗加寬
        dialog.minsize(280 if self.language == "ZH" else 340, 0)
        res_val = {"value": None}
        win_f = ("Microsoft JhengHei", 14)
        hint_f = ("Microsoft JhengHei", 11)
        hint_color = "#888888"

        ctk.CTkLabel(dialog,
                     text="請輸入時間格式：" if self.language == "ZH" else "Enter time format:",
                     font=win_f, anchor="w").pack(padx=30, pady=(20, 4), fill="x")
        entry = ctk.CTkEntry(dialog, font=win_f)
        entry.insert(0, self.excel_date_format)
        entry.pack(padx=30, fill="x")
        entry.focus_set()
        ctk.CTkLabel(dialog,
                     text="年:YY  月:M  日:DD  時:h  分:m  秒:ss  毫秒:SSS" if self.language == "ZH"
                     else "Year:YY  Month:M  Day:DD  Hour:h  Min:m  Sec:ss  Ms:SSS",
                     font=hint_f, text_color=hint_color, anchor="w").pack(padx=30, pady=(3, 0), fill="x")

        def on_ok():
            res_val["value"] = entry.get()
            dialog.destroy()

        def _focus_next_d(widgets, idx):
            widgets[(idx+1) % len(widgets)].focus_set()

        btn_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        btn_frame.pack(pady=16)
        btn_cancel = ctk.CTkButton(btn_frame, text=self.t("msg_cancel"), width=80,
                                   font=win_f, fg_color="gray", command=dialog.destroy)
        btn_cancel.pack(side="left", padx=10)
        btn_ok = ctk.CTkButton(btn_frame, text=self.t("msg_ok"), width=80,
                               font=win_f, command=on_ok)
        btn_ok.pack(side="left", padx=10)
        entry.bind("<Return>", lambda e: on_ok())

        self.root.wait_window(dialog)
        if res_val["value"] is not None:
            self.excel_date_format = res_val["value"]

    def change_utc_dialog(self):
        win = ctk.CTkToplevel(self.root)
        self.set_win_icon(win)
        win.title(self.t("tz_dialog_title"))
        win.attributes("-topmost", True)
        win.geometry(self.get_popup_pos(220))
        win.resizable(False, False)

        win_f = ("Microsoft JhengHei", 14)
        container = ctk.CTkFrame(win, fg_color="transparent")
        container.pack(padx=26, pady=18)

        current_var = tk.StringVar(value=self.format_utc_offset())
        row0 = ctk.CTkFrame(container, fg_color="transparent")
        row0.grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 4))
        ctk.CTkLabel(row0, text=f"{self.t('tz_dialog_current')}: ", font=("Microsoft JhengHei", 16, "bold")).pack(side="left")
        ctk.CTkLabel(row0, textvariable=current_var, font=("Microsoft JhengHei", 16, "bold")).pack(side="left", padx=(4, 0))

        ctk.CTkLabel(container, text=self.t("tz_dialog_prompt"), font=win_f, justify="left").grid(row=1, column=0, columnspan=2, sticky="w", pady=(0, 10))
        input_row = ctk.CTkFrame(container, fg_color="transparent")
        input_row.grid(row=2, column=0, columnspan=2, sticky="w")
        ctk.CTkLabel(input_row, text=self.t("tz_dialog_input"), font=win_f).pack(side="left", padx=(0, 8))
        input_var = tk.StringVar(value=self.format_utc_offset().replace("UTC", ""))
        entry_ctk = ctk.CTkEntry(input_row, width=160, textvariable=input_var, font=win_f)
        entry_ctk.pack(side="left")

        # 以分鐘整數保存，避免浮點誤差
        state = {"minutes": int(round(self.clamp_utc_offset(self.utc_offset) * 60))}

        def set_minutes(mins):
            mins = int(round(mins / 15) * 15)
            mins = max(-12 * 60, min(14 * 60, mins))
            state["minutes"] = mins
            current_var.set(self.format_utc_offset(mins / 60))
            input_var.set(self.format_utc_offset(mins / 60).replace("UTC", ""))

        def on_wheel(e):
            delta = 1 if e.delta > 0 else -1
            set_minutes(state["minutes"] + delta * 15)

        def apply_from_entry():
            try:
                v = self.parse_utc_offset_input(input_var.get())
                set_minutes(int(round(v * 60)))
            except:
                self.show_ctk_message(self.t("msg_hint"), self.t("err_format"))

        def on_ok():
            apply_from_entry()
            self.utc_offset = state["minutes"] / 60
            win.destroy()

        # bind wheel on window & entry widget
        win.bind("<MouseWheel>", on_wheel)
        entry_ctk.bind("<MouseWheel>", on_wheel)
        entry_ctk.bind("<Return>", lambda e: on_ok())
        entry_ctk.bind("<FocusOut>", lambda e: apply_from_entry())

        btn_frame = ctk.CTkFrame(win, fg_color="transparent")
        btn_frame.pack(pady=(0, 14))
        ctk.CTkButton(btn_frame, text=self.t("msg_cancel"), width=90, font=win_f, fg_color="gray", command=win.destroy).pack(side="left", padx=10)
        ctk.CTkButton(btn_frame, text=self.t("msg_ok"), width=90, font=win_f, command=on_ok).pack(side="left", padx=10)

    def change_text_color(self):
        picked = self.open_color_picker(self.t("pick_fg_title"), self.text_color)
        if picked:
            self.text_color = picked
            self._push_recent_color(picked)

    def change_bg_color(self):
        picked = self.open_color_picker(self.t("pick_bg_title"), self.bg_color)
        if picked:
            self.bg_color = picked
            self._push_recent_color(picked)

    def _push_recent_color(self, hex_color):
        recent = list(getattr(self, "recent_colors", []))
        hex_color = hex_color.upper()
        if hex_color in recent:
            recent.remove(hex_color)
        recent.insert(0, hex_color)
        self.recent_colors = recent[:10]

    def open_color_picker(self, title, initial_color="#FFFFFF"):
        import colorsys
        result = {"color": None}
        is_dark = self.appearance_mode == "dark"
        win_f = ("Microsoft JhengHei", 12)
        SQ = 180
        HB = 26

        def clamp(v, lo=0, hi=255): return max(lo, min(hi, int(v)))

        def hsv_to_hex(h, s, v):
            r, g, b = colorsys.hsv_to_rgb(h, s, v)
            return f"#{clamp(r*255):02X}{clamp(g*255):02X}{clamp(b*255):02X}"

        def hex_to_rgb(hx):
            hx = hx.lstrip("#")
            if len(hx) != 6: return None
            try: return int(hx[0:2],16), int(hx[2:4],16), int(hx[4:6],16)
            except: return None

        # 打開時直接使用傳入的當前顏色，不做白色替換
        # 把當前顏色加到歷程最前面（但不立即 save，等確定後才 push）
        rgb0 = hex_to_rgb(initial_color) or (255, 0, 0)
        h0, s0, v0 = colorsys.rgb_to_hsv(rgb0[0]/255, rgb0[1]/255, rgb0[2]/255)
        state = {"h": h0, "s": s0, "v": v0, "block": False}

        dialog = ctk.CTkToplevel(self.root)
        self.set_win_icon(dialog)
        dialog.title(title)
        dialog.attributes("-topmost", True)
        dialog.resizable(False, False)
        dialog.geometry(self.get_popup_pos(480))

        was_topmost = getattr(self, "is_topmost", True)

        # 強制讓 dialog 先渲染，再取正確背景色
        dialog.update_idletasks()
        try:
            raw_bg = dialog.cget("fg_color")
            if isinstance(raw_bg, (list, tuple)):
                # CTk fg_color tuple 格式是 (light_color, dark_color)
                bg_color = raw_bg[1] if is_dark else raw_bg[0]
            else:
                bg_color = raw_bg
        except:
            bg_color = "#2B2B2B" if is_dark else "#EBEBEB"

        pad = tk.Frame(dialog, bg=bg_color)
        pad.pack(padx=14, pady=12, fill="both", expand=True)

        top = tk.Frame(pad, bg=bg_color)
        top.pack()

        sq_border = tk.Frame(top, bg="#777777" if is_dark else "#888888",
                             padx=1, pady=1)
        sq_border.pack(side="left")
        sq = tk.Canvas(sq_border, width=SQ, height=SQ, highlightthickness=0,
                       cursor="crosshair")
        sq.pack()

        hbar_border = tk.Frame(top, bg="#777777" if is_dark else "#888888",
                               padx=1, pady=1)
        hbar_border.pack(side="left", padx=(8, 0))
        hbar = tk.Canvas(hbar_border, width=HB, height=SQ, highlightthickness=0,
                         cursor="sb_v_double_arrow")
        hbar.pack()

        mid = tk.Frame(pad, bg=bg_color)
        mid.pack(fill="x", pady=(10, 0))

        preview = tk.Canvas(mid, width=44, height=44, highlightthickness=1,
                            highlightbackground="#555" if is_dark else "#CCC")
        preview.pack(side="left")

        hex_var = tk.StringVar(value=hsv_to_hex(h0, s0, v0).upper())
        hex_entry = ctk.CTkEntry(mid, textvariable=hex_var, width=90, font=("Consolas", 13), justify="center")
        hex_entry.pack(side="left", padx=(10, 0))

        c_menu = self.get_menu_colors()
        c_picker = c_menu
        # 淺色模式用中灰（不突兀），深色模式用淡灰（清楚可見）
        icon_color = "#888888" if not is_dark else "#CCCCCC"
        hover_bg = c_menu["active_bg"]
        normal_bg = bg_color

        # CTkFrame 做圓角容器
        dropper_outer = ctk.CTkFrame(mid, width=40, height=40, corner_radius=8,
                                     fg_color=normal_bg, border_width=0)
        dropper_outer.pack(side="left", padx=(8, 0))
        dropper_outer.pack_propagate(False)

        drop_canvas = tk.Canvas(dropper_outer, width=36, height=36,
                                bg=normal_bg, highlightthickness=0, cursor="hand2")
        drop_canvas.place(relx=0.5, rely=0.5, anchor="center")

        def draw_dropper(hovered=False):
            bg = hover_bg if hovered else normal_bg
            drop_canvas.config(bg=bg)
            dropper_outer.configure(fg_color=bg)
            drop_canvas.delete("all")
            ic = icon_color
            # 管身：從右上斜向左下的粗線
            drop_canvas.create_line(27, 4, 13, 18, fill=ic, width=5, capstyle="round")
            # 管頂夾子（小矩形）
            drop_canvas.create_rectangle(24, 2, 31, 7, fill=ic, outline=ic)
            # 細管嘴
            drop_canvas.create_line(13, 18, 9, 22, fill=ic, width=2, capstyle="round")
            # 管嘴圓頭（替代水滴）
            drop_canvas.create_oval(5, 21, 13, 29, fill=ic, outline=ic)

        # Tooltip helper
        _tip_win = [None]
        def show_tip(widget, text):
            def on_enter(e):
                try:
                    if _tip_win[0]: _tip_win[0].destroy()
                except: pass
                tw = tk.Toplevel(dialog)
                tw.overrideredirect(True)
                tw.attributes("-topmost", True)
                lbl = tk.Label(tw, text=text, bg="#FFFFE0", fg="#333333",
                               font=("Microsoft JhengHei", 10), padx=6, pady=3,
                               relief="solid", bd=1)
                lbl.pack()
                x = widget.winfo_rootx() + widget.winfo_width() // 2
                y = widget.winfo_rooty() - 32
                tw.geometry(f"+{x}+{y}")
                _tip_win[0] = tw
            def on_leave(e):
                try:
                    if _tip_win[0]: _tip_win[0].destroy()
                    _tip_win[0] = None
                except: pass
            widget.bind("<Enter>", on_enter, add="+")
            widget.bind("<Leave>", on_leave, add="+")

        draw_dropper()
        drop_canvas.bind("<Enter>", lambda e: draw_dropper(True))
        drop_canvas.bind("<Leave>", lambda e: draw_dropper(False))
        drop_canvas.bind("<Button-1>", lambda e: do_eyedropper())
        show_tip(drop_canvas, "螢幕取色" if self.language == "ZH" else "Pick from screen")

        # 清除顏色歷程按鈕（垃圾桶）
        clear_outer = ctk.CTkFrame(mid, width=40, height=40, corner_radius=8,
                                   fg_color=normal_bg, border_width=0)
        clear_outer.pack(side="left", padx=(4, 0))
        clear_outer.pack_propagate(False)

        clear_canvas = tk.Canvas(clear_outer, width=36, height=36,
                                 bg=normal_bg, highlightthickness=0, cursor="hand2")
        clear_canvas.place(relx=0.5, rely=0.5, anchor="center")

        def draw_clear_btn(hovered=False):
            bg = hover_bg if hovered else normal_bg
            clear_canvas.config(bg=bg)
            clear_outer.configure(fg_color=bg)
            clear_canvas.delete("all")
            ic = icon_color
            # 垃圾桶蓋
            clear_canvas.create_rectangle(8, 8, 28, 11, fill=ic, outline=ic)
            clear_canvas.create_rectangle(13, 5, 23, 9, fill=ic, outline=ic)
            # 垃圾桶身
            clear_canvas.create_rectangle(10, 11, 26, 30, fill=ic, outline=ic)
            # 垃圾桶身內線（白色）
            body_bg = bg
            clear_canvas.create_line(15, 14, 15, 27, fill=body_bg, width=2)
            clear_canvas.create_line(18, 14, 18, 27, fill=body_bg, width=2)
            clear_canvas.create_line(21, 14, 21, 27, fill=body_bg, width=2)

        def on_clear_recent():
            self.recent_colors = []
            rebuild_recent()

        draw_clear_btn()
        clear_canvas.bind("<Enter>", lambda e: draw_clear_btn(True))
        clear_canvas.bind("<Leave>", lambda e: draw_clear_btn(False))
        clear_canvas.bind("<Button-1>", lambda e: on_clear_recent())
        show_tip(clear_canvas, "清除歷程" if self.language == "ZH" else "Clear history")

        sliders_frame = ctk.CTkFrame(pad, fg_color="transparent")
        sliders_frame.pack(fill="x", pady=(10, 0))
        sliders_frame.columnconfigure(1, weight=1)

        r_var = tk.IntVar()
        g_var = tk.IntVar()
        b_var = tk.IntVar()

        def make_rgb_row(label, var, color_fg, row):
            ctk.CTkLabel(sliders_frame, text=label, font=win_f, width=18, anchor="w",
                         text_color=color_fg).grid(row=row, column=0, sticky="w", padx=(0,6))

            def on_slide(val, _var=var):
                # 直接更新 var，繞過 trace 循環：先 block
                if state["block"]: return
                try:
                    new_v = int(float(val))
                    # 用 trace suspend 方式：直接改 state
                    state["block"] = True
                    _var.set(new_v)
                    ri, gi, bi = r_var.get(), g_var.get(), b_var.get()
                    h, s, v = colorsys.rgb_to_hsv(ri/255, gi/255, bi/255)
                    state["h"] = h; state["s"] = s; state["v"] = v
                    refresh_all(defer_sq=True)
                    state["block"] = False
                except: state["block"] = False

            sl = ctk.CTkSlider(sliders_frame, from_=0, to=255, variable=var,
                               height=16, command=on_slide,
                               button_color=color_fg, button_hover_color=color_fg,
                               progress_color=color_fg)
            sl.grid(row=row, column=1, sticky="ew", padx=(0,6))
            num = ctk.CTkEntry(sliders_frame, width=46, font=("Consolas", 12), justify="center")
            num.grid(row=row, column=2)
            return sl, num

        sl_r, num_r = make_rgb_row("R", r_var, "#E05555", 0)
        sl_g, num_g = make_rgb_row("G", g_var, "#55AA55", 1)
        sl_b, num_b = make_rgb_row("B", b_var, "#5588DD", 2)

        recent_frame = tk.Frame(pad, bg=bg_color)
        recent_frame.pack(fill="x", pady=(10, 0))
        lbl_color = "#888888"
        lbl_font = ("Microsoft JhengHei", 9)

        # 整排橫向容器
        swatches_row = tk.Frame(recent_frame, bg=bg_color)
        swatches_row.pack(fill="x")
        recent_swatches = []

        def rebuild_recent():
            for w in recent_swatches:
                try: w.destroy()
                except: pass
            recent_swatches.clear()

            cur = initial_color.upper()
            history = [hx.upper() for hx in getattr(self, "recent_colors", [])[:10]
                       if hx.upper() != cur][:9]

            # ── 當前色塊 ──
            cur_col = tk.Frame(swatches_row, bg=bg_color)
            cur_col.pack(side="left", padx=(0, 4))
            recent_swatches.append(cur_col)

            sw_cur = tk.Canvas(cur_col, width=22, height=22,
                               highlightthickness=1,
                               highlightbackground="#555" if is_dark else "#CCC",
                               cursor="hand2")
            try: sw_cur.config(bg=cur)
            except: pass
            sw_cur.pack()
            sw_cur.bind("<Button-1>", lambda e, c=cur: apply_hex(c))
            tk.Label(cur_col, text="當前" if self.language == "ZH" else "Now",
                     bg=bg_color, fg=lbl_color, font=lbl_font).pack()

            # ── 垂直分隔線 ──
            sep_v = tk.Frame(swatches_row, bg="#4A4A4A" if is_dark else "#BBBBBB",
                             width=1)
            sep_v.pack(side="left", fill="y", padx=(2, 4), pady=2)
            recent_swatches.append(sep_v)

            if not history:
                return

            # ── 歷程區塊（標籤 + 色塊） ──
            hist_col = tk.Frame(swatches_row, bg=bg_color)
            hist_col.pack(side="left")
            recent_swatches.append(hist_col)

            # 歷程標籤放在第一排色塊下方對齊
            hist_top = tk.Frame(hist_col, bg=bg_color)
            hist_top.pack(fill="x")

            for hx in history:
                cell = tk.Frame(hist_top, bg=bg_color)
                cell.pack(side="left", padx=2)
                sw_c = tk.Canvas(cell, width=22, height=22,
                                 highlightthickness=1,
                                 highlightbackground="#555" if is_dark else "#CCC",
                                 cursor="hand2")
                try: sw_c.config(bg=hx)
                except: continue
                sw_c.pack()
                sw_c.bind("<Button-1>", lambda e, c=hx: apply_hex(c))
                recent_swatches.append(sw_c)

            tk.Label(hist_col,
                     text="顏色歷程" if self.language == "ZH" else "History",
                     bg=bg_color, fg=lbl_color, font=lbl_font).pack(anchor="w")

        def _close_dialog():
            if was_topmost:
                self.root.attributes("-topmost", True)
            dialog.destroy()

        btns = ctk.CTkFrame(pad, fg_color="transparent")
        btns.pack(pady=(14, 0))
        ctk.CTkButton(btns, text=self.t("msg_cancel"), width=90, font=win_f,
                      fg_color="gray", command=_close_dialog).pack(side="left", padx=8)
        ctk.CTkButton(btns, text=self.t("msg_ok"), width=90, font=win_f,
                      command=lambda: [
                          result.update({"color": hsv_to_hex(state["h"], state["s"], state["v"])}),
                          self._push_recent_color(initial_color),
                          _close_dialog()
                      ]).pack(side="left", padx=8)

        # ── 繪製：用 numpy 加速 HSV 方塊 ─────────────────
        _sq_img = [None]
        _hb_img = [None]
        _sq_after = [None]

        def draw_sq_now():
            try:
                from PIL import Image, ImageTk
                try:
                    import numpy as np
                    h = state["h"]
                    s_arr = np.linspace(0, 1, SQ)
                    v_arr = np.linspace(1, 0, SQ)
                    S, V = np.meshgrid(s_arr, v_arr)
                    H = np.full((SQ, SQ), h)
                    # HSV -> RGB via numpy
                    hi = (H * 6).astype(int)
                    f = H * 6 - hi
                    p = V * (1 - S)
                    q = V * (1 - f * S)
                    t = V * (1 - (1 - f) * S)
                    hi = hi % 6
                    R = np.select([hi==0,hi==1,hi==2,hi==3,hi==4,hi==5],[V,q,p,p,t,V])
                    G = np.select([hi==0,hi==1,hi==2,hi==3,hi==4,hi==5],[t,V,V,q,p,p])
                    B = np.select([hi==0,hi==1,hi==2,hi==3,hi==4,hi==5],[p,p,t,V,V,q])
                    rgb = np.stack([R,G,B], axis=2)
                    rgb = (rgb * 255).clip(0,255).astype(np.uint8)
                    img = Image.fromarray(rgb, "RGB")
                except ImportError:
                    img = Image.new("RGB", (SQ, SQ))
                    px = img.load()
                    h = state["h"]
                    for xi in range(SQ):
                        s = xi / (SQ-1)
                        for yi in range(SQ):
                            v = 1.0 - yi/(SQ-1)
                            r,g,b = colorsys.hsv_to_rgb(h,s,v)
                            px[xi,yi] = (clamp(r*255),clamp(g*255),clamp(b*255))
                tkimg = ImageTk.PhotoImage(img)
                _sq_img[0] = tkimg
                sq.delete("all")
                sq.create_image(0, 0, anchor="nw", image=tkimg)
            except Exception:
                pass

        def draw_sq(defer=False):
            if defer:
                if _sq_after[0]: dialog.after_cancel(_sq_after[0])
                _sq_after[0] = dialog.after(80, draw_sq_now)
            else:
                draw_sq_now()

        def draw_hbar():
            try:
                from PIL import Image, ImageTk
                img = Image.new("RGB", (HB, SQ))
                px = img.load()
                for yi in range(SQ):
                    h = yi / (SQ-1)
                    r,g,b = colorsys.hsv_to_rgb(h, 1.0, 1.0)
                    for xi in range(HB):
                        px[xi, yi] = (clamp(r*255), clamp(g*255), clamp(b*255))
                tkimg = ImageTk.PhotoImage(img)
                _hb_img[0] = tkimg
                hbar.delete("all")
                hbar.create_image(0, 0, anchor="nw", image=tkimg)
                hy = int(state["h"] * (SQ-1))
                # 圓角矩形指示線，顏色跟選單 active_bg 一致
                ac = c_picker["active_bg"]
                r_ind = 3
                hbar.create_rectangle(1, max(0, hy-r_ind), HB-1, min(SQ-1, hy+r_ind),
                                      outline=ac, fill="", width=2)
            except: pass

        def update_preview():
            hx = hsv_to_hex(state["h"], state["s"], state["v"])
            try: preview.config(bg=hx)
            except: pass
            hex_var.set(hx.upper())

        def update_rgb_from_state():
            r,g,b = colorsys.hsv_to_rgb(state["h"], state["s"], state["v"])
            ri,gi,bi = clamp(r*255), clamp(g*255), clamp(b*255)
            r_var.set(ri); g_var.set(gi); b_var.set(bi)
            num_r.delete(0,"end"); num_r.insert(0,str(ri))
            num_g.delete(0,"end"); num_g.insert(0,str(gi))
            num_b.delete(0,"end"); num_b.insert(0,str(bi))

        def refresh_all(redraw_sq=True, defer_sq=False):
            if redraw_sq:
                draw_sq(defer=defer_sq)
            draw_hbar()
            update_preview()
            update_rgb_from_state()

        def sq_pick(e):
            if state["block"]: return
            state["s"] = max(0.0, min(1.0, e.x/(SQ-1)))
            state["v"] = max(0.0, min(1.0, 1.0-e.y/(SQ-1)))
            hx = hsv_to_hex(state["h"], state["s"], state["v"])
            try:
                preview.config(bg=hx)
            except: pass
            hex_var.set(hx.upper())
            update_rgb_from_state()

        sq.bind("<Button-1>", sq_pick)
        sq.bind("<B1-Motion>", sq_pick)
        sq.bind("<ButtonRelease-1>", lambda e: draw_sq(defer=False))

        def hbar_pick(e):
            if state["block"]: return
            state["block"] = True
            state["h"] = max(0.0, min(1.0, e.y/(SQ-1)))
            # s=0 或 v=0 時，顏色固定是灰/黑，色相無效；自動設為有色模式
            if state["s"] < 0.05:
                state["s"] = 1.0
            if state["v"] < 0.05:
                state["v"] = 1.0
            update_preview()
            update_rgb_from_state()
            draw_hbar()
            draw_sq(defer=True)
            state["block"] = False

        hbar.bind("<Button-1>", hbar_pick)
        hbar.bind("<B1-Motion>", hbar_pick)
        hbar.bind("<ButtonRelease-1>", lambda e: draw_sq(defer=False))

        def on_rgb_slider(*_):
            if state["block"]: return
            state["block"] = True
            ri,gi,bi = r_var.get(), g_var.get(), b_var.get()
            h,s,v = colorsys.rgb_to_hsv(ri/255, gi/255, bi/255)
            state["h"]=h; state["s"]=s; state["v"]=v
            refresh_all(defer_sq=True)
            state["block"] = False

        r_var.trace_add("write", on_rgb_slider)
        g_var.trace_add("write", on_rgb_slider)
        b_var.trace_add("write", on_rgb_slider)

        def on_num_entry(num_widget, var, event=None):
            try:
                v = clamp(int(num_widget.get()))
                var.set(v)
            except: pass

        num_r.bind("<Return>", lambda e: on_num_entry(num_r, r_var))
        num_g.bind("<Return>", lambda e: on_num_entry(num_g, g_var))
        num_b.bind("<Return>", lambda e: on_num_entry(num_b, b_var))
        num_r.bind("<FocusOut>", lambda e: on_num_entry(num_r, r_var))
        num_g.bind("<FocusOut>", lambda e: on_num_entry(num_g, g_var))
        num_b.bind("<FocusOut>", lambda e: on_num_entry(num_b, b_var))

        def apply_hex(hx=None):
            val = hx or hex_var.get().strip()
            rgb = hex_to_rgb(val)
            if not rgb: return
            h,s,v = colorsys.rgb_to_hsv(rgb[0]/255, rgb[1]/255, rgb[2]/255)
            state["h"]=h; state["s"]=s; state["v"]=v
            refresh_all()

        hex_entry.bind("<Return>", lambda e: apply_hex())
        hex_entry.bind("<FocusOut>", lambda e: apply_hex())

        def do_eyedropper():
            dialog.withdraw()
            self.root.after(120, _run_eyedropper)

        def _run_eyedropper():
            picked = self.pick_color_from_screen()
            dialog.deiconify()
            dialog.lift()
            if picked:
                apply_hex(picked)

        rebuild_recent()
        # 立即繪製，不延遲，確保打開時顏色正確
        dialog.update_idletasks()
        draw_hbar()
        refresh_all(redraw_sq=True)
        self.root.wait_window(dialog)
        # 確保置頂恢復（X 關閉時 protocol 已處理，這裡作為保險）
        if was_topmost:
            try: self.root.attributes("-topmost", True)
            except: pass
        return result["color"]


    def pick_color_from_screen(self):

        user32 = ctypes.windll.user32
        gdi32 = ctypes.windll.gdi32

        def get_pixel_hex(x, y):
            hdc = user32.GetDC(0)
            color = gdi32.GetPixel(hdc, int(x), int(y))
            user32.ReleaseDC(0, hdc)
            if color == -1: return None
            r = color & 0xFF; g = (color >> 8) & 0xFF; b = (color >> 16) & 0xFF
            return f"#{r:02X}{g:02X}{b:02X}"

        _pt = ctypes.wintypes.POINT()
        user32.GetCursorPos(ctypes.byref(_pt))
        start_x, start_y = _pt.x, _pt.y

        try:
            import mss as _mss
            _sct = _mss.mss()
            has_mss = True
        except ImportError:
            has_mss = False

        try:
            from PIL import Image, ImageTk
            has_pil = True
        except ImportError:
            has_pil = False

        if not (has_mss or has_pil):
            return None

        result = {"color": None}
        _state = {"running": True, "after_id": None}
        _cur = [start_x, start_y]
        _last = [-999, -999]
        _tkimg_ref = [None]

        # 取得所有螢幕的總範圍
        try:
            SM_XVIRTUALSCREEN  = 76
            SM_YVIRTUALSCREEN  = 77
            SM_CXVIRTUALSCREEN = 78
            SM_CYVIRTUALSCREEN = 79
            vx = ctypes.windll.user32.GetSystemMetrics(SM_XVIRTUALSCREEN)
            vy = ctypes.windll.user32.GetSystemMetrics(SM_YVIRTUALSCREEN)
            vw = ctypes.windll.user32.GetSystemMetrics(SM_CXVIRTUALSCREEN)
            vh = ctypes.windll.user32.GetSystemMetrics(SM_CYVIRTUALSCREEN)
        except:
            vx, vy, vw, vh = 0, 0, user32.GetSystemMetrics(0), user32.GetSystemMetrics(1)

        picker = tk.Toplevel(self.root)
        picker.overrideredirect(True)
        picker.attributes("-topmost", True)
        picker.attributes("-alpha", 0.01)
        picker.configure(bg="black")
        picker.geometry(f"{vw}x{vh}+{vx}+{vy}")
        picker.config(cursor="none")

        HUD_SIZE = 140
        ZOOM_PX = 13
        HUD_OFFSET = 24

        hud = tk.Toplevel(self.root)
        hud.overrideredirect(True)
        hud.attributes("-topmost", True)
        hud.configure(bg="#111111")

        # 游標箭頭 label（顯示在 HUD 外側靠近游標的角落）
        ARROW = "▶"  # 會依方位旋轉文字替代
        cursor_lbl = tk.Label(hud, text="✛", bg="#111111", fg="#FFFFFF",
                              font=("Consolas", 9), padx=0, pady=0)

        canvas = tk.Canvas(hud, width=HUD_SIZE, height=HUD_SIZE, bg="#111111",
                           highlightthickness=2, highlightbackground="#555555")
        canvas.pack()
        color_label = tk.Label(hud, text="#------", bg="#111111", fg="#FFFFFF",
                               font=("Consolas", 11, "bold"), pady=3)
        color_label.pack(fill="x")
        hud.geometry(f"+{start_x+HUD_OFFSET}+{start_y+HUD_OFFSET}")

        # 游標指示視窗：透明背景 + 十字，緊貼游標尖端
        TRANS_KEY = "#010203"
        CROSS_SIZE = 40
        CROSS_HOLE = 18
        arrow_win = tk.Toplevel(self.root)
        arrow_win.overrideredirect(True)
        arrow_win.attributes("-topmost", True)
        arrow_win.configure(bg=TRANS_KEY)
        try:
            arrow_win.attributes("-transparentcolor", TRANS_KEY)
        except: pass
        arrow_canvas = tk.Canvas(arrow_win, width=CROSS_SIZE, height=CROSS_SIZE,
                                 bg=TRANS_KEY, highlightthickness=0)
        arrow_canvas.pack()

        def draw_cross():
            arrow_canvas.delete("all")
            c2 = CROSS_SIZE // 2
            h = CROSS_HOLE // 2
            lw = 3
            # 垂直線上半
            arrow_canvas.create_line(c2, 0, c2, c2-h, fill="black", width=lw+2, capstyle="butt")
            arrow_canvas.create_line(c2, 0, c2, c2-h, fill="white", width=lw, capstyle="butt")
            # 垂直線下半
            arrow_canvas.create_line(c2, c2+h, c2, CROSS_SIZE, fill="black", width=lw+2, capstyle="butt")
            arrow_canvas.create_line(c2, c2+h, c2, CROSS_SIZE, fill="white", width=lw, capstyle="butt")
            # 水平線左半
            arrow_canvas.create_line(0, c2, c2-h, c2, fill="black", width=lw+2, capstyle="butt")
            arrow_canvas.create_line(0, c2, c2-h, c2, fill="white", width=lw, capstyle="butt")
            # 水平線右半
            arrow_canvas.create_line(c2+h, c2, CROSS_SIZE, c2, fill="black", width=lw+2, capstyle="butt")
            arrow_canvas.create_line(c2+h, c2, CROSS_SIZE, c2, fill="white", width=lw, capstyle="butt")

        draw_cross()
        arrow_win.geometry(f"{CROSS_SIZE}x{CROSS_SIZE}+{start_x - CROSS_SIZE//2}+{start_y - CROSS_SIZE//2}")

        def update_loop():
            if not _state["running"]:
                return
            mx, my = _cur[0], _cur[1]
            if mx != _last[0] or my != _last[1]:
                _last[0] = mx; _last[1] = my
                try:
                    half = ZOOM_PX // 2
                    if has_mss:
                        mon = {"left": mx-half, "top": my-half,
                               "width": ZOOM_PX, "height": ZOOM_PX}
                        sshot = _sct.grab(mon)
                        img = Image.frombytes("RGB", sshot.size, sshot.bgra, "raw", "BGRX")
                    else:
                        from PIL import ImageGrab
                        img = ImageGrab.grab(bbox=(mx-half, my-half, mx+half+1, my+half+1), all_screens=True)
                    img = img.resize((HUD_SIZE, HUD_SIZE), Image.NEAREST)
                    tkimg = ImageTk.PhotoImage(img)
                    _tkimg_ref[0] = tkimg
                    canvas.delete("all")
                    canvas.create_image(0, 0, anchor="nw", image=tkimg)
                    m = HUD_SIZE // 2
                    # 只畫中心小方框，不畫十字（十字在外部視窗）
                    canvas.create_rectangle(m-3, m-3, m+3, m+3, outline="black", width=2)
                    canvas.create_rectangle(m-3, m-3, m+3, m+3, outline="white", width=1)
                    color_label.config(text=get_pixel_hex(mx, my) or "#------")

                    # 四方位自動切換
                    sw2 = vw + vx; sh2 = vh + vy
                    hud.update_idletasks()
                    hud_w = hud.winfo_reqwidth()
                    hud_h = hud.winfo_reqheight()
                    right_ok  = mx + HUD_OFFSET + hud_w <= sw2
                    bottom_ok = my + HUD_OFFSET + hud_h <= sh2

                    if right_ok and bottom_ok:
                        hx2, hy2 = mx + HUD_OFFSET, my + HUD_OFFSET
                    elif right_ok and not bottom_ok:
                        hx2, hy2 = mx + HUD_OFFSET, my - HUD_OFFSET - hud_h
                    elif not right_ok and bottom_ok:
                        hx2, hy2 = mx - HUD_OFFSET - hud_w, my + HUD_OFFSET
                    else:
                        hx2, hy2 = mx - HUD_OFFSET - hud_w, my - HUD_OFFSET - hud_h

                    hud.geometry(f"+{hx2}+{hy2}")
                    # 十字永遠置中在游標位置
                    ax = mx - CROSS_SIZE // 2
                    ay = my - CROSS_SIZE // 2
                    arrow_win.geometry(f"{CROSS_SIZE}x{CROSS_SIZE}+{ax}+{ay}")
                except Exception:
                    pass
            _state["after_id"] = picker.after(16, update_loop)

        def on_move(e):
            _cur[0] = e.x_root; _cur[1] = e.y_root

        def cleanup():
            _state["running"] = False
            if _state["after_id"]:
                try: picker.after_cancel(_state["after_id"])
                except: pass
            try:
                if has_mss: _sct.close()
            except: pass
            for w in (arrow_win, hud, picker):
                try:
                    if w.winfo_exists(): w.destroy()
                except: pass

        def on_pick(e):
            result["color"] = get_pixel_hex(e.x_root, e.y_root)
            cleanup()

        def on_cancel(e=None):
            cleanup()

        picker.bind("<Motion>", on_move)
        picker.bind("<Button-1>", on_pick)
        picker.bind("<Button-3>", on_cancel)
        picker.bind("<Escape>", on_cancel)
        picker.grab_set()
        _state["after_id"] = picker.after(16, update_loop)
        self.root.wait_window(picker)
        return result["color"]

    def on_mouse_wheel(self, e):
        new_val = max(0.1, min(1.0, self.bg_alpha + (0.05 if e.delta > 0 else -0.05)))
        self.bg_alpha = new_val
        # 滾輪整體調整：text_alpha 和 bg_fill_alpha 同步
        self.text_alpha = new_val
        self.bg_fill_alpha = new_val
        self.apply_alpha_settings()
        self._show_alpha_toast(new_val)

    def _show_alpha_toast(self, val):
        pct = int(round(val * 100))
        txt = f"{pct}%"
        if hasattr(self, "_toast_win") and self._toast_win:
            try:
                if self._toast_win.winfo_exists():
                    self._toast_label.config(text=txt)
                    if hasattr(self, "_toast_after"):
                        self.root.after_cancel(self._toast_after)
                    self._toast_after = self.root.after(1200, self._hide_toast)
                    self._reposition_toast()
                    return
            except: pass
        tw = tk.Toplevel(self.root)
        tw.overrideredirect(True)
        tw.attributes("-topmost", True)
        tw.attributes("-alpha", 0.75)
        tw.configure(bg="#111111")
        lbl = tk.Label(tw, text=txt, bg="#111111", fg="#FFFFFF",
                       font=("Consolas", 15, "bold"), padx=10, pady=5)
        lbl.pack()
        self._toast_win = tw
        self._toast_label = lbl
        tw.update_idletasks()
        self._reposition_toast()
        self._toast_after = self.root.after(1200, self._hide_toast)

    def _reposition_toast(self):
        try:
            if not (hasattr(self, "_toast_win") and self._toast_win and self._toast_win.winfo_exists()):
                return
            tw = self._toast_win
            tw.update_idletasks()
            rx = self.root.winfo_x()
            ry = self.root.winfo_y()
            rw = self.root.winfo_width()
            rh = self.root.winfo_height()
            tw_w = tw.winfo_reqwidth()
            tw_h = tw.winfo_reqheight()
            x = rx + rw - tw_w - 4
            y = ry + rh - tw_h - 4
            tw.geometry(f"+{x}+{y}")
        except: pass

    def _hide_toast(self):
        if hasattr(self, "_toast_win") and self._toast_win:
            try:
                if self._toast_win.winfo_exists():
                    self._toast_win.destroy()
            except: pass
            self._toast_win = None

    def start_move(self, e): self.x, self.y = e.x, e.y
    def do_move(self, e):
        self.root.geometry(f"+{self.root.winfo_x()+(e.x-self.x)}+{self.root.winfo_y()+(e.y-self.y)}")
        self._reposition_toast()

if __name__ == "__main__":
    root = ctk.CTk()
    app = FloatingClock(root)
    root.mainloop()