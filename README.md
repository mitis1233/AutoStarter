## AutoStarter 自動啟動器

一鍵自動化啟動流程：啟動應用程式&設定音訊裝置工具

## 截圖

![介面](https://raw.githubusercontent.com/mitis1233/AutoStarter/refs/heads/main/Picture.png)

## 功能介紹

新增程式：可啟動多個應用程式 可指定啟動參數

啟動最小化： 打勾將會最小化來啟動應用程式

強制最小化：監控連帶開啟的子進程 確保所有相關窗口都被最小化 假如最小化沒效可使用

新增延遲：可插入啟動延遲 可自定義延遲秒數

音訊裝置：可切換、停用 撥放設備及錄製設備

自定義流程： 自訂從上而下的啟動順序 可按鈕&右鍵選單操控

檔案關聯： 首次啟動須點選註冊檔案關聯 關聯後可直接開啟.autostart檔案運行 除非你更換AutoStarter位置才要重新關聯

匯入設定檔： 從以儲存.autostart匯入設定檔 也可以滑鼠拖放.autostart檔案匯入

輸出設定檔： 將您的自動化流程儲存為.autostart設定檔

## 使用說明

首次啟動須點選註冊檔案關聯

1. 設定好自動化啟動流程

2. 點選輸出設定檔按鈕儲存.autostart設定檔

3. 開啟.autostart設定檔即可啟動自定義的自動化啟動流程

## 下載 Download

[獨立式](https://github.com/mitis1233/AutoStarter/releases/latest/download/AutoStarter_win-x64_Full.zip)

[相依 需下載Net9](https://github.com/mitis1233/AutoStarter/releases/latest/download/AutoStarter_win-x64.zip)




## ✨ 以下AI詳細介紹

### 🚀 應用程式管理
- **新增程式**：支持啟動多個應用程式，可指定自定義啟動參數
- **最小化啟動**：支持普通最小化和強制最小化兩種模式
  - 普通最小化：標準 Windows 最小化
  - 強制最小化：監控子進程，確保所有相關窗口都被最小化
- **參數自動清理**：自動移除啟動參數前後的空格，避免執行錯誤

### ⏱️ 流程控制
- **新增延遲**：在操作序列中插入自定義延遲時間（秒數可調）
- **順序管理**：通過上移/下移按鈕或右鍵菜單自訂執行順序

### 🔊 音訊裝置管理
- **切換音訊設備**：快速切換播放設備和錄製設備
- **停用音訊設備**：臨時停用指定的音訊裝置
- 支持多個音訊端點（Console、Multimedia、Communications）

### 💾 設定檔管理
- **輸出設定檔**：將自動化流程保存為 `.autostart` 設定檔
- **匯入設定檔**：支持三種匯入方式
  - 按鈕匯入：點擊「匯入設定檔」按鈕
  - 拖放匯入：直接拖動 `.autostart` 檔案到 GUI（v1.4 新增）
  - 檔案關聯：雙擊 `.autostart` 檔案自動打開並執行

### 📋 檔案關聯
- **一鍵註冊**：首次使用時點擊「註冊檔案關聯」
- **便捷執行**：註冊後可直接雙擊 `.autostart` 檔案運行
- **靈活移除**：若更換 AutoStarter 位置，點擊「移除檔案關聯」後重新註冊

---

## 🎯 使用流程

### 快速開始

1. **首次啟動**
   - 打開 AutoStarter
   - 點擊「註冊檔案關聯」（僅需一次）

2. **建立自動化流程**
   - 點擊「新增程式」選擇要啟動的應用程式
   - 設置啟動參數（可選）
   - 勾選「最小化」或「強制最小化」（可選）
   - 根據需要添加延遲或音訊設備操作
   - 使用上移/下移調整執行順序

3. **保存設定檔**
   - 點擊「輸出設定檔」按鈕
   - 選擇保存位置和檔案名
   - 點擊「保存」

4. **執行自動化流程**
   - 方式 A：雙擊保存的 `.autostart` 檔案
   - 方式 B：在 AutoStarter 中點擊「匯入設定檔」後執行
   - 方式 C：拖動 `.autostart` 檔案到 AutoStarter GUI 中

### 高級用法

#### 複雜流程示例
```
1. 啟動 exe（帶參數）- 強制最小化
2. 延遲 2 秒
3. 切換音訊設備到「耳機」
4. 啟動 exe（帶參數）- 普通最小化
5. 延遲 1 秒
6. 停用「麥克風」
```

## 🔧 技術架構

### 技術棧
- **框架**：WPF (Windows Presentation Foundation)
- **語言**：C# 12 (.NET 9.0)
- **依賴**：
  - NAudio 2.2.1 - 音訊設備管理
  - System.Management 10.0.0 - 進程管理和 WMI 查詢

### 核心組件

| 組件 | 功能 |
|------|------|
| `MainWindow.xaml(.cs)` | 主 GUI 界面、事件處理、拖放功能 |
| `ActionItem.cs` | 操作項數據模型，支持 MVVM 綁定 |
| `App.xaml.cs` | 應用啟動邏輯、.autostart 檔案執行引擎 |
| `AudioDeviceSelectorWindow.xaml(.cs)` | 音訊設備選擇對話框 |
| `EditActionWindow.xaml(.cs)` | 操作項編輯對話框 |
| `CoreAudio/` | 音訊設備管理核心庫 |

### 操作類型
```csharp
enum ActionType
{
    LaunchApplication,      // 啟動應用程式
    Delay,                  // 延遲等待
    SetAudioDevice,         // 切換音訊設備
    DisableAudioDevice      // 停用音訊設備
}
```

### 設定檔格式 (.autostart)
```json
[
  {
    "MinimizeWindow": true,
    "ForceMinimizeWindow": false,
    "Type": "LaunchApplication",
    "FilePath": "C:\\Program Files\\App\\app.exe",
    "Arguments": "--fullscreen -screen-width 2560",
    "DelaySeconds": 0,
    "AudioDeviceId": null,
    "AudioDeviceName": null
  },
  {
    "MinimizeWindow": false,
    "ForceMinimizeWindow": false,
    "Type": "Delay",
    "FilePath": "",
    "Arguments": "",
    "DelaySeconds": 2,
    "AudioDeviceId": null,
    "AudioDeviceName": null
  }
]
```

---

## 🆕 v1.4 更新內容

### 新增功能
- ✨ **拖放匯入**：支持直接拖動 `.autostart` 檔案到 GUI 進行匯入
- 🧹 **參數自動清理**：匯出時自動移除啟動參數的前後空格
- 📦 **依賴更新**：System.Management 升級至 v10.0.0

### 改進
- 優化了拖放事件的視覺反饋
- 改進了參數驗證邏輯
- 增強了錯誤處理機制

### 已知特性
- 支持子進程最小化監控（v1.3+）
- 互斥的最小化選項（普通/強制）
- 完整的音訊設備管理
