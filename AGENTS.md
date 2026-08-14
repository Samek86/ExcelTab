<!-- Generated: 2026-08-10 | Updated: 2026-08-10 -->

# ExcelTab

## Purpose

Windows デスクトップ向けの Excel ブック・タブスイッチャー。Microsoft Excel の複数ブックをブラウザのようなタブバーで切り替え、お気に入りジャンプ・ショートカット・YAML 設定を提供する WPF ユーティリティ。トレイ／ツールウィンドウとして画面下部に常駐し、COM 経由で実行中の Excel と連動する。

## Key Files

| File | Description |
|------|-------------|
| `ExcelTab.sln` | Visual Studio ソリューション（単一プロジェクト） |
| `説明書.xlsx` | ユーザー向け操作説明書 |
| `テストファイル.xlsx` | 動作確認用のテスト用ワークブック |

## Subdirectories

| Directory | Purpose |
|-----------|---------|
| `ExcelTab/` | WPF アプリケーション本体のソース・プロジェクト (see `ExcelTab/AGENTS.md`) |

**除外・生成物（編集対象外）:** `ExcelTab/bin/`（ビルド出力・実行時設定）、`ExcelTab/obj/`（中間生成）、`.vs/`（IDE キャッシュ）

## Architecture Overview

```
App.xaml.cs
  └─ 二重起動制御 → ApplicationSetting.Load() → MainWindow
MainWindow
  ├─ Excel COM (Marshal.GetActiveObject / Interop)
  ├─ ROTManager（開いているローカル Workbook 列挙）
  ├─ KeyHookManager（グローバルショートカット）
  ├─ TabItemControl 一覧（タブ UI）
  └─ VIEW/*（お気に入り・切替・ダイアログ）
設定: ExcelTabSetting.yaml（exe 同ディレクトリ、YamlDotNet）
```

## Tech Stack

| Area | Technology |
|------|------------|
| Language | C# |
| UI | WPF (XAML + code-behind) |
| Runtime | .NET Framework **4.7.2** |
| Build | Visual Studio / MSBuild（クラシック非 SDK スタイル `.csproj`） |
| Excel | [NetOfficeFw.Excel](https://www.nuget.org/packages/NetOfficeFw.Excel/) 1.9.10（`NetOffice.ExcelApi`） |
| Packaging | Costura.Fody（依存 DLL を exe に埋め込み） |
| Settings | YamlDotNet → `ExcelTabSetting.yaml` |
| Tray | Hardcodet.NotifyIcon.Wpf |
| Icons | MahApps.Metro.IconPacks.BoxIcons |
| Input | MouseKeyboardActivityMonitor + カスタム KeyboardHook |

## For AI Agents

### Working In This Directory

- ソース変更は主に `ExcelTab/` 配下。ルートの `.xlsx` はドキュメント／テストデータであり、通常はコード変更の対象外。
- README / 自動テスト / CI は現状なし。変更後は Visual Studio または MSBuild でビルドし、Excel 起動状態での手動確認が前提。
- UI 文言・コメントは主に日本語。新機能のユーザー向け文字列も日本語に合わせる。
- `bin/`・`obj/` はコミット・編集しない。実行時設定 `ExcelTabSetting.yaml` は `bin/Release` 等に生成される。

### Build

```powershell
# Visual Studio で ExcelTab.sln を開くか、MSBuild 例:
msbuild ExcelTab.sln /p:Configuration=Release
```

- ビルドには **Microsoft Excel（Office COM）がインストールされた Windows** 環境が望ましい（Interop 参照）。
- 成果物: `ExcelTab\bin\Release\ExcelTab.exe`（Costura により単一 exe 化）

### Testing Requirements

- 専用テストプロジェクトなし。手動確認を推奨:
  1. Excel で複数ブックを開く
  2. ExcelTab 起動 → タブ表示・切替
  3. ショートカット（Alt+NumPad / Alt+Left/Right 等）
  4. お気に入り登録・ジャンプ
  5. 設定 YAML の保存・再読込
- ルートの `テストファイル.xlsx` を検証用に利用可能。

### Common Patterns

- 名前空間: `ExcelTab`, `ExcelTab.CLASSES`, `ExcelTab.VIEW`, `ExcelTab.ITEM`, `ExcelTab.Extension`
- Excel API エイリアス: `using Excel = Microsoft.Office.Interop.Excel;`
- 設定は静的 `App.Setting`（`ApplicationSetting`）経由
- 共通処理は静的 `Common` クラスに集約

## Dependencies

### Internal

- すべて `ExcelTab/` 単一プロジェクト内。ソリューションに他プロジェクトなし。

### External

- .NET Framework 4.7.2 / WPF
- Office Excel COM Interop
- NuGet: Costura.Fody, Fody, Hardcodet.NotifyIcon.Wpf, MahApps.Metro.IconPacks.BoxIcons, MouseKeyboardActivityMonitor, YamlDotNet, WPF.ColorPicker, System.ValueTuple

## Constraints & Risks

- **COM 寿命**: Excel オブジェクトは `Marshal.ReleaseComObject` / `FinalReleaseComObject` に注意。リークは Excel プロセス残留につながる。
- **STA / UI スレッド**: WPF Dispatcher 経由で UI 更新（`Common.Invoke` 等）。
- **グローバルフック**: キーボードフックは他アプリに影響しうる。解放漏れに注意。
- **二重起動**: 既存インスタンスを Kill する挙動あり（`App.xaml.cs`）。変更時は副作用を意識する。
- **対象 OS**: Windows 専用。macOS/Linux 非対応。

<!-- MANUAL: Any manually added notes below this line are preserved on regeneration -->
