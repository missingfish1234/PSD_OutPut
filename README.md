# PSD Export Pipeline

Version `1.1.46` / Build `v36` / Stamp `2026-03-24-02`

Photoshop UXP 面板，支援批次匯出 PNG、`metadata/layout.json`、Spine、Unity 6.0 / 6.3 匯入包，以及 Cocos Creator 3.8.8 直接 Prefab 匯入包。

## 功能

- 批次匯出 PNG
- 輸出 `metadata/layout.json`
- 輸出 `metadata/export_report.json` 與 `metadata/export_report.txt`
- 輸出 Spine：`skeleton.json`、`.atlas`、`spine/images/*.png`
- 輸出 Unity 6.0 Prefab 建置包
- 輸出 Unity 6.3 最小匯入包與 `.unitypackage` staging payload
- 輸出 Cocos Creator 3.8.8 最小 direct prefab 匯入包
- 圖層順序：`auto` / `psd` / `reverse`
- 資料夾驅動命名：`主資料夾_001` / `主資料夾_語系_001`
- 匯出後自檢：缺檔、重名、全透明 PNG、文字圖層渲染風險

## 精簡輸出規則

當 `Prefab 目標平台` 選擇下列項目時，工具會改用精簡輸出：

- `Unity 6.3`
  - 不保留根目錄 `images/`
  - 不保留 `metadata/layout.json`
  - 只保留建 prefab 與封 `.unitypackage` 必需檔
- `Cocos Creator 3.8.8`
  - 不保留根目錄 `images/`
  - 不保留 `metadata/layout.json`
  - 不再輸出舊版 `resources/layout/extensions` fallback
  - 只保留 direct prefab 必需檔

`metadata/export_report.json` 與 `metadata/export_report.txt` 仍會保留。

## 命名規則

圖片檔名改為依資料夾命名，不再直接用圖層名稱：

- 一般資料夾：
  - `TotalWin/...` 會輸出成 `Total_001`、`Total_002`
- 多國語系資料夾：
  - `TotalWin/JA/...` 會輸出成 `TotalWin_JA_001`
  - `TotalWin/CHT/...` 會輸出成 `TotalWin_CHT_001`
  - `TotalWin/CHS/...` 會輸出成 `TotalWin_CHS_001`
  - `TotalWin/EN/...` 會輸出成 `TotalWin_EN_001`

規則細節：

- 主資料夾名稱取第一個有效群組名稱
- 若有語系子資料夾，會改用完整主資料夾名加語系碼
- 序號固定為三位數：`001`、`002`、`003`

## 使用方式

1. 在 UDT 載入 [manifest.json](/f:/pstesttool/Test-7ld4kd/manifest.json)
2. 在 Photoshop 開啟 PSD
3. 開啟面板 `PSD Export Pipeline`
4. 選擇輸出資料夾
5. 設定輸出模式與 Prefab 平台
6. 按 `重新掃描 PSD`
7. 按 `開始匯出`

## 輸出結構

### 一般模式 / Unity 6.0

```text
<export-folder>/
  images/
    *.png
  metadata/
    layout.json
    export_report.json
    export_report.txt
  spine/
    skeleton.json
    *.atlas
    images/
      *.png
  engine/
    unity6_0/
      images/
        *.png
      layout_unity6.json
      Editor/
        PsdUnity6PrefabAutoBuilder.cs
```

### Unity 6.3 精簡模式

```text
<export-folder>/
  metadata/
    export_report.json
    export_report.txt
  engine/
    unity6_3/
      Assets/
        PSDExport/
          <PSD名稱>/
            layout_unity6_3.json
            layout_unity6_3.json.meta
            Images/
              *.png
              *.png.meta
            Editor/
              PsdUnity63PrefabAutoBuilder.cs
              PsdUnity63PrefabAutoBuilder.cs.meta
```

### Cocos Creator 3.8.8 精簡模式

```text
<export-folder>/
  metadata/
    export_report.json
    export_report.txt
  engine/
    cocos3_8_8/
      <PSD名稱>_PSDAssets/
        Prefab/
          <PSD名稱>_PSD.prefab
          <PSD名稱>_PSD.prefab.meta
        Texture/
          psd_export/
            *.png
            *.png.meta
      <PSD名稱>_PSDAssets.meta
```

## Unity 6.3

直接匯入 Unity 專案時，複製：

- `engine/unity6_3/Assets/PSDExport/<PSD名稱>/`

到 Unity 專案 `Assets/` 之後，`PsdUnity63PrefabAutoBuilder.cs` 會自動掃描：

- `Assets/PSDExport/<PSD名稱>/layout_unity6_3.json`

並在同資料夾建立：

- `<PSD名稱>_PSD.prefab`

若要封成 `.unitypackage`：

Unity 6.3 現在只輸出可直接導入 Unity 專案的資料夾。
- 保留 `engine/unity6_3/Assets/PSDExport/<PSD名稱>/`
- 不再輸出 `unitypackage_payload`、`unitypackage_manifest.json`、`Build_Unity6_3_UnityPackage.ps1`

## Cocos Creator 3.8.8

直接匯入時，複製：

- `engine/cocos3_8_8/<PSD名稱>_PSDAssets/`
- `engine/cocos3_8_8/<PSD名稱>_PSDAssets.meta`

到 Cocos 專案 `assets/`。

Prefab 位置：

- `assets/<PSD名稱>_PSDAssets/Prefab/<PSD名稱>_PSD.prefab`

## 自檢報告

`metadata/export_report.json` 與 `metadata/export_report.txt` 目前會檢查：

- 缺檔
- 重名
- rename 警告
- 路徑警告
- Spine 排序規則測試
- Unity 6.3 最小匯出關鍵檔案
- Cocos 3.8.8 direct prefab 關鍵檔案
- 匯出後全透明 PNG
- 全透明文字圖層匯出風險

## 主要檔案

- [manifest.json](/f:/pstesttool/Test-7ld4kd/manifest.json)
- [index.html](/f:/pstesttool/Test-7ld4kd/index.html)
- [index.js](/f:/pstesttool/Test-7ld4kd/index.js)
