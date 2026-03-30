# 🕐 Floating-Clock

A lightweight always-on-top floating clock for Windows with a modern UI and full customization.

輕量、現代化的 Windows 懸浮時鐘，支援透明背景與高度自訂樣式。

---

## ✨ Features / 功能特色

- 懸浮於所有視窗之上 / Always on top
- 透明背景，融入桌面 / Transparent background
- 深色 / 淺色主題 / Dark & Light mode
- 12H / 24H、星期、時區、毫秒顯示 / Multiple time formats
- 顯示 Unix 時間戳 / Unix timestamp display
- 自訂文字與背景顏色（含螢幕取色器）/ Custom colors with eyedropper
- 自訂時間顯示格式 / Custom time format
- 日期與時間戳互轉工具 / Datetime ↔ Timestamp converter
- 滾輪調整透明度 / Scroll to adjust opacity
- Ctrl + 滾輪縮放字體 / Ctrl + Scroll to zoom font
- 繁體中文 / 英文介面 / Traditional Chinese & English UI

---

## 🚀 Getting Started / 開始使用

### 直接執行 / Run EXE

前往 [Releases](../../releases) 下載最新版 `Floating-Clock.exe`，與 `app.ico` 放在同一資料夾後直接執行。

Download the latest `Floating-Clock.exe` from [Releases](../../releases), place it in the same folder as `app.ico`, and run.

### 原始碼執行 / Run from Source

```bash
pip install customtkinter Pillow mss
python Floating-Clock.py
```

---

## 🖱️ 使用說明 / How to Use

| 操作 / Action | 功能 / Function |
|---|---|
| 左鍵拖曳 / Left drag | 移動時鐘 / Move clock |
| 左鍵雙擊 / Left double-click | 關閉時鐘 / Close clock |
| 右鍵點擊 / Right-click | 開啟選單 / Open menu |
| 滾輪上下 / Scroll | 調整透明度 / Adjust opacity ±5% |
| Ctrl + 滾輪 / Ctrl + Scroll | 縮放字體 / Zoom font |

### 選單功能 / Menu

- **格式設定** — 切換時間格式、時區、星期、時間戳顯示
- **樣式設定** — 變更文字與背景顏色
- **設定** — 切換語言、深淺色主題
- **時間戳轉換器** — 日期與 Unix 時間戳互轉
- **恢復預設樣式** — 重設所有外觀設定

---

## ⚙️ 設定儲存位置 / Config Location

```
%LOCALAPPDATA%\Floating-Clock\clock_config.json
```

---

## 🛠️ 系統需求 / Requirements

- Windows 10 / 11
- Python 3.10+（原始碼執行 / source only）