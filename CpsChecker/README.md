# CpsChecker

![CpsChecker Icon](path/to/your/icon.png) <!-- 請將此路徑替換為你的程式圖標實際路徑 -->

CpsChecker 是一款用於分析字幕檔（.srt/.ass/.ssa）中每秒字元數（CPS）的工具，可協助你找出過快的字幕行，提升觀影體驗。

[英文版 README](README_en.md) | [程式流程圖](flowchart.png) <!-- 請將這些路徑替換為實際的英文版README和流程圖檔案路徑 -->

## 功能特色
- 支援自動偵測字幕編碼（整合 Ude）
- 支援命令列參數：靜默模式、詳細模式、輸出報告
- 支援 .NET Framework 4.8.1 編譯

## 使用方式
CpsChecker.exe -f <字幕檔或資料夾> [-o <報告檔>] [-v]
### 參數說明
- `-f <檔案或資料夾>`：指定字幕檔或目錄 (支援 .srt .ass .ssa)
- `-o <輸出檔>`：指定報告輸出位置 (預設為 report.txt)
- `-v`：顯示詳細處理步驟
- `-h`：顯示說明

### 錯誤碼
- `0` = 成功
- `1` = 找不到字幕檔
- `2` = 檔案解析錯誤
- `3` = 指令參數錯誤

## 範例# 分析指定字幕檔案，輸出報告到 default_report.txt
CpsChecker.exe -f input.srt

# 分析指定目錄下的所有字幕檔案，輸出報告到 custom_report.txt
CpsChecker.exe -f /path/to/subtitles -o custom_report.txt

# 分析指定字幕檔案，顯示詳細處理步驟
CpsChecker.exe -f input.srt -v
## 授權
本專案採用 MIT 授權，詳細內容請參考 [LICENSE](LICENSE)。 <!-- 請將此路徑替換為實際的授權檔案路徑 -->
