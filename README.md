# RAIN-DelayBurst

Windows 網路測試小工具（PySide6 + WinDivert），可對指定程式做上下行「擠壓回放 / 丟包」模擬。

## 快速開始

1. 安裝套件
```powershell
pip install -r requirements.txt
```

2. 確認檔案同目錄
- `RAIN-DelayBurst.py`
- `WinDivert.dll`
- `WinDivert64.sys`

3. 用系統管理員執行
```powershell
python RAIN-DelayBurst.py
```

## 使用方式

1. 點 `選擇目標` 選擇 `.exe`
2. 點 `綁定熱鍵`，按下要綁定的鍵
3. 設定上行 / 下行參數（可各自啟用）
4. 用熱鍵或 `切換效果` 啟用，`結束效果` 停止
5. 可 `保存設定` / `載入設定`

## 模式說明

- `擠壓回放`：封包先扣押，效果結束後依設定回放  
- `丟包模式`：效果期間按機率直接丟包

## 原理簡述

1. 透過 `tasklist` / `netstat` 找出目標程式的 PID 與使用中的本機埠。  
2. 用 WinDivert 建立上下行過濾規則並攔截封包。  
3. `擠壓回放` 模式會先暫存封包，效果結束後依抖動、帶寬與丟包參數回放。  
4. `丟包模式` 會在效果期間依丟包率直接丟棄封包。  

## 使用的庫

- WinDivert: <https://github.com/basil00/WinDivert>

## 常見問題

- `找不到 WinDivert DLL`：把 `WinDivert.dll` 放在程式同資料夾
- `請用系統管理員身分執行`：以系統管理員啟動
