# Excel「重新計算」巨集原始碼

- `Module1.bas`：巨集原始碼（`RecalcAll`），`{{中文}}` 會在建置時轉成 `ChrW(...)`，讓程式碼不受 Windows 語系影響。
- `Sheet1_events.bas`：主工作表的事件（J2 選圖、K2～M2 勾選、B6:C7 上下限變更時重畫）。
- `Module2.bas`：「亮度量測」工作表（A19 = No.）的重新計算與亮度分布圖／推移圖重畫（`RecalcLum`、`LumAuto`），圖上文字與顯示設定讀取 `_cie_params` 的 O／P 欄。
- `ThisWorkbook_events.bas`：活頁簿事件，「亮度量測」的亮度值（C 欄）或規格（B6:D6）變更時自動重畫。
- `build_vba.py`：依 MS-OVBA／MS-CFB 規格產生只含原始碼的 `vbaProject.bin`，Excel 開啟時會自行編譯。

修改巨集後重新產生：

```
python3 build_vba.py Module1.bas vbaProject.bin Sheet1_events.bas --module Module2.bas --workbook ThisWorkbook_events.bas
```

再把 `vbaProject.bin` 轉成 base64，取代 `index.html` 裡 `VBA_PROJECT_B64` 的內容。

- `check_explicit.py`：依 Excel `Option Explicit` 規則檢查每個程序的變數是否都有宣告。
