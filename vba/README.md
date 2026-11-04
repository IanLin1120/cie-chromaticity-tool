# Excel「重新計算」巨集原始碼

- `Module1.bas`：巨集原始碼（`RecalcAll`），`{{中文}}` 會在建置時轉成 `ChrW(...)`，讓程式碼不受 Windows 語系影響。
- `build_vba.py`：依 MS-OVBA／MS-CFB 規格產生只含原始碼的 `vbaProject.bin`，Excel 開啟時會自行編譯。

修改巨集後重新產生：

```
python3 build_vba.py Module1.bas vbaProject.bin
```

再把 `vbaProject.bin` 轉成 base64，取代 `index.html` 裡 `VBA_PROJECT_B64` 的內容。
