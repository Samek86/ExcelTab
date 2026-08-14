<!-- Parent: ../AGENTS.md -->
<!-- Generated: 2026-08-10 | Updated: 2026-08-10 -->

# VIEW

## Purpose

メインタブバー以外の WPF ウィンドウとユーザーコントロール。お気に入り、ブック切替 UI、ダイアログ、個別タブ項目など。各画面は `.xaml` + `.xaml.cs` のペア。

## Key Files

| File | Description |
|------|-------------|
| `TabItemControl.xaml` / `.xaml.cs` | タブ1枚分の UI。Workbook 参照・プレビュー画像・色。ドラッグは 8px 以上移動してから開始 |
| `FavoriteWindow.xaml` / `.xaml.cs` | お気に入り一覧ウィンドウ |
| `FavoriteItemControl.xaml` / `.xaml.cs` | お気に入り1件の表示・操作コントロール |
| `ExcelSwitchingWindow.xaml` / `.xaml.cs` | Excel ブック切替用オーバーレイ／ウィンドウ |
| `SwitchingControl.xaml` / `.xaml.cs` | 切替 UI の部品コントロール |
| `PromptDialog.xaml` / `.xaml.cs` | 入力・確認用ダイアログ |
| `AlertWindow.xaml` / `.xaml.cs` | 通知アラート。`Alert` クラス + `AlertStateEnum` |

## Subdirectories

なし。

## For AI Agents

### Working In This Directory

- 新規画面は **XAML と code-behind を同時追加**し、`ExcelTab.csproj` の `<Page>` と `<Compile DependentUpon=...>` を更新する。
- 既定 namespace は `ExcelTab.VIEW`（`TabItemControl` 含む）。
- スタイルは `STYLE/` の ResourceDictionary またはインライン。アプリ全体スタイルは `App.xaml` の MergedDictionaries 経由。
- Excel COM や `MainWindow.TabList` など静的状態を触る場合、UI スレッド制約を守る。
- ウィンドウはメインと同様に Topmost / 位置計算が絡むことがある。`Common.GetWorkingArea` / `Magnification` と整合させる。

### Testing Requirements

- 各 Window の表示・閉じる・Esc／ボタン操作
- お気に入り: 登録 → 一覧 → セル／シートジャンプ → 設定 YAML への永続化
- 切替 UI: 複数ブック時の選択と Excel 前面化
- Alert: 成功／失敗表示の色・自動クローズがあれば時間確認

### Common Patterns

- partial class + `InitializeComponent()`
- UserControl は親（MainWindow / FavoriteWindow 等）から生成・リスト管理
- お気に入りデータモデルは `ExcelTab.ITEM.FavoriteExcel`

## Dependencies

### Internal

- `ExcelTab`（MainWindow, App.Setting, Common, TabItemControl）
- `ExcelTab.CLASSES`（フック・Win32 が必要な場合）
- `ExcelTab.ITEM`（FavoriteExcel）
- `ExcelTab/STYLE/`（見た目）

### External

- WPF (`System.Windows.*`)
- Excel Interop（ブック／セル操作を行うコントロール）
- MahApps IconPacks / ColorPicker（使用箇所による）

<!-- MANUAL: Any manually added notes below this line are preserved on regeneration -->
