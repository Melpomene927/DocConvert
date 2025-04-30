# Word 文件格式轉換工具

這個工具組提供了兩種轉換功能：
1. 將 Microsoft Word 舊格式文件 (.doc) 批量轉換為新格式 (.docx)
2. 將 Microsoft Word 文件 (.docx) 批量轉換為純文字檔案 (.txt)

## 系統需求

- Windows 作業系統
- Microsoft Word（任何支援 .docx 格式的版本）
- PowerShell 3.0 或以上版本（Windows 7以上版本系統內建）

## 成功執行步驟

### 方法1: 使用批處理檔案（推薦）

1. 打開DOC目錄
2. 雙擊 `ConvertDocs.bat`
3. 從選單中選擇一個選項：
   - **DOC 轉 DOCX 選項:**
     - 選項 1: 標準轉換 (保留原始檔案)
     - 選項 2: 轉換並刪除原始檔案
     - 選項 3: 不建立備份的轉換
     - 選項 4: 不建立備份並刪除原始檔案
   - **DOCX 轉 TXT 選項:**
     - 選項 5: 將 DOCX 轉換為 TXT (保留原始檔案)
     - 選項 6: 將 DOCX 轉換為 TXT 並刪除原始檔案
   - **其他選項:**
     - 選項 7: 自定義參數
     - 選項 8: 離開

4. 如果選擇選項7，按提示輸入以下參數：
   - 轉換類型：選擇1進行DOC轉DOCX，或選擇2進行DOCX轉TXT
   - 來源目錄：輸入包含要處理檔案的目錄路徑（可以是絕對路徑或相對路徑）
   - 是否刪除原始檔案：輸入y/n
   - 是否建立備份：輸入y/n
   - 是否使用安靜模式：輸入y/n

### 方法2: 直接使用PowerShell腳本

1. 打開PowerShell或命令提示字元視窗
2. 導航到DOC目錄
3. 執行以下命令:

   ```powershell
   # DOC 轉 DOCX - 基本用法（轉換當前目錄中的檔案）
   powershell -ExecutionPolicy Bypass -File ConvertDocToDocx.ps1
   
   # DOC 轉 DOCX - 指定不同的源目錄（使用相對路徑）
   powershell -ExecutionPolicy Bypass -File ConvertDocToDocx.ps1 -SourceDir "File Description_PY"
   
   # DOCX 轉 TXT - 基本用法（轉換當前目錄中的檔案）
   powershell -ExecutionPolicy Bypass -File ConvertDocxToTxt.ps1
   
   # DOCX 轉 TXT - 指定不同的源目錄
   powershell -ExecutionPolicy Bypass -File ConvertDocxToTxt.ps1 -SourceDir "FileLayout_PY"
   
   # 轉換後刪除原始檔案
   powershell -ExecutionPolicy Bypass -File ConvertDocxToTxt.ps1 -DeleteOriginal
   
   # 不建立備份
   powershell -ExecutionPolicy Bypass -File ConvertDocxToTxt.ps1 -Backup:$false
   ```

## 處理特殊情況

### 處理目錄連結

Windows目錄連結（Directory Junction）需要特別處理：

1. 對於檔案連結，如DOC目錄下的子目錄是連結到其他位置，請直接指定實際目標位置：
   ```powershell
   # 假設DOC\FileLayout_PY是指向D:\ARTHUR\ARTHPY\Doc\FileLayout的連結
   powershell -ExecutionPolicy Bypass -File ConvertDocToDocx.ps1 -SourceDir "D:\ARTHUR\ARTHPY\Doc\FileLayout"
   ```

### 中文檔案路徑處理

腳本已經最佳化，可以正確處理包含中文字符的檔案名稱和路徑。確保您的PowerShell終端使用正確的字元編碼（建議使用UTF-8）。

## 執行結果

成功執行腳本後，您將看到：

1. 轉換進度顯示在命令視窗中
2. 轉換完成後的統計資訊
3. 生成的日誌檔案（根據轉換類型生成 DocConversion.log 或 DocxToTxtConversion.log）
4. 如果啟用備份，將創建一個日期時間戳名稱的備份目錄，位於DOC目錄中

## 常見問題排解

1. **找不到檔案**
   - 確認指定的路徑是否正確
   - 檢查該路徑下是否確實有要轉換的檔案（注意副檔名大小寫）
   - 如果是目錄連結，嘗試直接指定目標目錄路徑

2. **參數錯誤**
   - 確保執行腳本時使用正確的參數語法
   - 確保開關參數（如`-DeleteOriginal`）使用正確的語法

3. **轉換失敗**
   - 確保Microsoft Word已正確安裝
   - 檢查是否有檔案被鎖定或需要管理員權限
   - 查看日誌檔案了解詳細錯誤信息

## 使用範例

以下是一些常見使用場景：

1. **標準轉換場景**
   - 使用批處理檔案選項1或5，或直接執行無參數腳本

2. **大量檔案轉換**
   - 推薦啟用備份功能以確保資料安全
   - 可以使用安靜模式減少輸出

3. **連結目錄處理**
   - 使用絕對路徑指定實際目標目錄

4. **空間有限的環境**
   - 使用`-Backup:$false`選項跳過備份
   - 或使用`-DeleteOriginal`選項在轉換後刪除原始檔案

## 詳細參數說明

- `-SourceDir`：指定要處理的目錄，預設為 "."（當前目錄）
- `-DeleteOriginal`：轉換後刪除原始檔案
- `-Backup`：在執行轉換前建立備份，預設為 $true
- `-Silent`：減少輸出訊息

## 備份功能

啟用備份功能時，腳本會在執行前建立一個名為 "DocBackup_日期時間" 或 "DocxBackup_日期時間" 的備份目錄，並保存所有原始檔案的副本。

## 日誌檔案

轉換過程的詳細記錄會保存在對應的日誌檔案中，包含：
- 處理時間
- 成功轉換的檔案
- 跳過的檔案
- 失敗的檔案
- 備份位置（如果啟用）

## 常見問題

1. **執行腳本時出現安全錯誤**
   
   如果您第一次在系統上運行 PowerShell 腳本，可能需要修改 PowerShell 的執行策略。批處理檔案已包含必要的 `-ExecutionPolicy Bypass` 參數，但如果直接執行 .ps1 檔案，可能需要手動設定。

2. **Word 授權對話框出現**
   
   如果在執行過程中看到 Word 彈出授權請求對話框，請接受/啟用它。這是由於腳本使用 Word COM 物件自動處理檔案所致。

3. **某些檔案轉換失敗**
   
   可能有些檔案是受密碼保護或損壞的。這些檔案會被記錄在日誌中，您可能需要手動處理它們。

4. **DOCX 轉 TXT 後文字格式問題**

   由於純文字格式不支援複雜的格式設定，轉換後的 .txt 檔案僅包含純文字內容，所有格式設定（如粗體、斜體、顏色等）都將丟失。表格和圖片也會以簡化的方式轉換或可能丟失。

## 授權

此工具組僅供內部使用，供 NetGUIBA 專案相關人員使用。