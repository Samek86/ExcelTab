# ExcelTab

Windows 向けの Excel ブック・タブスイッチャーです。開いているブックを画面下部のタブで切り替え、お気に入りジャンプとショートカットを提供します。

## 要件

- Windows
- .NET Framework 4.7.2
- Microsoft Excel

## ビルド

Visual Studio で `ExcelTab.sln` を開くか:

```powershell
msbuild ExcelTab.sln /p:Configuration=Release
```

成果物: `ExcelTab\bin\Release\ExcelTab.exe`

## 技術

- WPF / C#
- Excel 自動化: [NetOfficeFw.Excel](https://www.nuget.org/packages/NetOfficeFw.Excel/) 1.9.10
- 設定: `ExcelTabSetting.yaml`（exe と同じフォルダ）
