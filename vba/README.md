# Excel「重新計算」巨集原始碼

- `Module1.bas`：巨集原始碼（`RecalcAll`），`{{中文}}` 會在建置時轉成 `ChrW(...)`，讓程式碼不受 Windows 語系影響。
- `Sheet1_events.bas`：主工作表的事件（J2 選圖、K2～M2 勾選、B6:C7 上下限變更時重畫）。
- `build_vba.py`：依 MS-OVBA／MS-CFB 規格產生只含原始碼的 `vbaProject.bin`，Excel 開啟時會自行編譯。

修改巨集後重新產生：

```
python3 build_vba.py Module1.bas vbaProject.bin Sheet1_events.bas
```

再把 `vbaProject.bin` 轉成 base64，取代 `index.html` 裡 `VBA_PROJECT_B64` 的內容。
