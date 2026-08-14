<!-- Parent: ../AGENTS.md -->
<!-- Generated: 2026-08-10 | Updated: 2026-08-10 -->

# ExcelTab (project)

## Purpose

WPF アプリケーションのプロジェクトルート。エントリポイント（`App`）、メインタブウィンドウ（`MainWindow`）、プロジェクト定義・Fody 設定、および機能別サブフォルダをまとめる。

## Key Files

| File | Description |
|------|-------------|
| `ExcelTab.csproj` | クラシック csproj。COM 参照・PackageReference・XAML/CS 一覧 |
| `App.xaml` / `App.xaml.cs` | アプリ起動。二重起動チェック、設定ロード、TEMP 作成、MainWindow 表示。ResourceDictionary マージ |
| `MainWindow.xaml` / `MainWindow.xaml.cs` | 画面下部のタブバー本体。Excel 監視、タブ構築、Win32 スタイル（ツールウィンドウ・Topmost・キャプチャ除外） |
| `App.config` | ランタイム設定（.NET 4.7.2） |
| `FodyWeavers.xml` | Costura による依存埋め込み設定 |
| `ExcelTab2.ico` | アプリケーションアイコン |

## Subdirectories

| Directory | Purpose |
|-----------|---------|
| `CLASSES/` | コアロジック（設定、COM/ROT、フック、Win32、画面、更新）(see `CLASSES/AGENTS.md`) |
| `VIEW/` | ウィンドウ・ユーザーコントロール（タブ以外の UI）(see `VIEW/AGENTS.md`) |
| `STYLE/` | 共有 XAML ResourceDictionary (see `STYLE/AGENTS.md`) |
| `ITEM/` | ドメインモデル・列挙型 (see `ITEM/AGENTS.md`) |
| `EXTENTION/` | 拡張メソッド（フォルダ名の綴りは Extention）(see `EXTENTION/AGENTS.md`) |
| `Properties/` | AssemblyInfo・Resources・manifest・Settings (see `Properties/AGENTS.md`) |
| `bin/` | ビルド出力（編集禁止） |
| `obj/` | 中間生成（編集禁止） |

## Namespaces

| Namespace | Location |
|-----------|----------|
| `ExcelTab` | App, MainWindow, Common, ApplicationSetting |
| `ExcelTab.CLASSES` | ROTManager, KeyHookManager, ExcelHelper, FavoriteHelper, CaptureHelper, ComHelper, AppLog 等 |
| `ExcelTab.VIEW` | 各 Window / UserControl（TabItemControl 含む） |
| `ExcelTab.ITEM` | FavoriteExcel, WinkEnum |
| `ExcelTab.Extension` | StringExtension |
| `ExcelTab.Properties` | 生成リソース・Settings |

## Runtime Paths

| Path | Role |
|------|------|
| `{BaseDirectory}/ExcelTabSetting.yaml` | ユーザー設定（タブ色・名前・ショートカット・お気に入り・IgnoreFiles） |
| `{BaseDirectory}/TEMP/` | 一時ファイル（お気に入りサムネ等） |

`BaseDirectory` は通常 `bin\Debug` または `bin\Release`。

## For AI Agents

### Working In This Directory

- 新規 `.cs` / `.xaml` を追加したら **`ExcelTab.csproj` の `<Compile>` / `<Page>` にも登録**する（SDK スタイルではないため自動 incl されない）。
- XAML と code-behind はペアで管理。`x:Class` と partial class 名を一致させる。
- `App.xaml` の MergedDictionaries に載っていない STYLE はアプリ全体では未使用（コメントアウト済みのものあり）。
- `TabItemControl` の namespace は `ExcelTab.VIEW`。

### Startup Flow

1. 同名プロセスが複数なら他を `Kill`
2. `ApplicationSetting.Load()` → なければ生成して Save
3. `Common.CreateDirectory(TempFolderPath)`
4. `new MainWindow().Show()`（UpdateManager はコメントアウト）

### MainWindow Responsibilities

- 作業領域下部への配置・DPI/倍率（`Common.Magnification`）
- `GetExcelApp()` で `Excel.Application` 取得
- `SetTabSp()` / `ExcelMonitor` でタブ同期
- `KeyHookManager` でショートカット
- コンテキストメニュー（スタートアップ登録、Excel 全終了、終了など）

### Testing Requirements

- Release/Debug ビルド後、Excel を起動した状態で exe を実行
- 設定変更後は `ExcelTabSetting.yaml` の内容と再起動後の反映を確認

### Common Patterns

- 静的状態: `MainWindow.oExcelApp`, `MainWindow.TabList`, `App.Setting`
- Win32: `Win32Helper` の P/Invoke を `using static` で利用
- UI スレッド: `Common.Invoke` / Dispatcher

## Dependencies

### Internal

- 全サブフォルダの型に依存。中心は `CLASSES` と `VIEW`。

### External

- PackageReference（csproj）: Costura.Fody 3.3.3, Fody 4.2.1, Hardcodet.NotifyIcon.Wpf, MahApps.Metro.IconPacks.BoxIcons, MouseKeyboardActivityMonitor, YamlDotNet 11.2.1, WPF.ColorPicker, System.ValueTuple
- COMReference: Microsoft.Office.Core, Microsoft.Office.Interop.Excel, VBIDE

<!-- MANUAL: Any manually added notes below this line are preserved on regeneration -->
