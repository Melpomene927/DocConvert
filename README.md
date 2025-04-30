# DocConvert

文件轉換工具集，用於處理各種文件格式的轉換需求。

## 環境需求

- Windows 10 或更新版本
- PowerShell 7.0 或更新版本
- .NET Framework 4.7.2 或更新版本

## 專案結構

```
DocConvert/
├── Documents/        # 文件相關資料
├── Scripts/         # 腳本檔案
│   ├── PowerShell/  # PowerShell 腳本
│   └── Batch/       # Batch 腳本
└── .specstory/      # 規格說明檔案
```

## 安裝說明

1. 複製專案到本機
```powershell
git clone [repository-url]
cd DocConvert
```

2. 確認 PowerShell 執行權限
```powershell
Get-ExecutionPolicy
# 如果不是 RemoteSigned 或 Unrestricted，請執行：
Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
```

## 使用說明

### PowerShell 腳本

PowerShell 腳本位於 `Scripts/PowerShell/` 目錄下，使用前請確認執行權限。

### Batch 腳本

Batch 腳本位於 `Scripts/Batch/` 目錄下，可直接執行。

## 注意事項

1. 執行腳本前請確認檔案編碼為 ANSI，以確保中文字符正確顯示
2. 建議使用系統管理員權限執行腳本
3. 請參考各腳本檔案開頭的說明來了解使用方式

## 貢獻指南

1. 腳本開發規範
   - PowerShell 腳本使用 `.ps1` 副檔名
   - Batch 腳本使用 `.bat` 或 `.cmd` 副檔名
   - 遵循專案的程式碼風格指南
   
2. 提交規範
   - 提交訊息請使用繁體中文
   - 技術名詞保持英文原文

## 授權

本專案採用 MIT 授權條款 - 詳見 [LICENSE](LICENSE) 檔案 