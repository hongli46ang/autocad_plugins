# AutoTableFromDataLink.bundle

C# / AutoCAD .NET 外掛，用現有 Data Link 自動建立表格。

## 指令

- `AUTOTABLEDATALINK`
- 短指令：`AUTOTABLEDL`

## 流程

1. 選資料連結編號。
2. 選表格插入位置。
3. 自動插入 Data Link 表格，並套用：
   - 儲存格寬度：`2025`
   - 儲存格高度：`486`
   - 文字高度：`200`

## 建置

此專案預設給 AutoCAD 2021 使用，目標框架是 `.NET Framework 4.8`（`net48`）。
請先確認電腦有安裝 `.NET Framework 4.8 Developer Pack` 或 Visual Studio / Build Tools 的 `.NET desktop development` 工作負載。

目前的 `.NET SDK 6.0.428` 可以拿來執行建置命令；但若缺少 4.8 targeting pack，會無法編譯 `net48`。可以先確認 SDK：

```powershell
dotnet --list-sdks
```

```powershell
cd AutoTableFromDataLink.bundle\Contents\Source
dotnet build -c Release -p:AcadInstallDir="C:\Program Files\Autodesk\AutoCAD 2021"
```

建置完成後，DLL 會輸出到：

```text
AutoTableFromDataLink.bundle\Contents\Windows\AutoTableFromDataLink.dll
```

如果 AutoCAD 安裝在其他版本或路徑，把 `AcadInstallDir` 改成你的 AutoCAD 目錄。

> 注意：AutoCAD 2021 使用 .NET Framework 4.8。AutoCAD 2025+ 則改用 .NET 8，不能共用同一個 DLL。

## 安裝

把整個 `AutoTableFromDataLink.bundle` 放到：

```text
%AppData%\Autodesk\ApplicationPlugins
```

或建立資料夾連結：

```powershell
New-Item -ItemType SymbolicLink -Path "$env:AppData\Autodesk\ApplicationPlugins\AutoTableFromDataLink.bundle" -Target "C:\Users\user\路徑\autocad_plugins\AutoTableFromDataLink.bundle"
```

如果目前圖面還沒有 Data Link，可先用既有 `AUTODATALINKSHEETS` 建立 Excel 分頁資料連結。
