<!-- Parent: ../AGENTS.md -->
<!-- Generated: 2026-08-10 | Updated: 2026-08-10 -->

# CLASSES

## Purpose

Excel 連携・入力フック・Win32・設定・画面・更新など、UI 以外のコアロジックを置くディレクトリ。アプリの振る舞いの中心。

## Key Files

| File | Description |
|------|-------------|
| `ApplicationSetting.cs` | YAML 設定モデル。タブ色/名前、ショートカット文字列、お気に入り、IgnoreFiles。Load/Save（YamlDotNet）。namespace は `ExcelTab` |
| `Common.cs` | 共有状態と UI 補助（パス、色、Dispatcher、DPI、Wink）。Excel/キャプチャ/COM は Helper へ分割。namespace は `ExcelTab` |
| `AppLog.cs` | exe 同ディレクトリの `ExcelTab.log` と Debug 出力 |
| `ComHelper.cs` | COM 参照の Release |
| `CaptureHelper.cs` | ウィンドウキャプチャと Bitmap / ImageSource 変換 |
| `ExcelHelper.cs` | アクティブブック判定とセル編集（コピー・連番・背景など） |
| `FavoriteHelper.cs` | お気に入り登録・ジャンプ |
| `ROTManager.cs` | Running Object Table から開いている Excel Workbook を列挙。IgnoreFiles でアドイン等を除外 |
| `KeyHookManager.cs` | キーボードフックの管理。ショートカット文字列変換・タブ切替などのキー処理 |
| `KeyboardHook.cs` | 低レベル／ライブラリ連携のキーボードフック実装 |
| `Win32Helper.cs` | P/Invoke（ウィンドウスタイル、Topmost、表示親和性、SetWindowPos 等） |
| `ScreenManager.cs` | マルチモニタ／画面関連の補助 |
| `UpdateManager.cs` | 更新チェック（現状 App 側では呼び出しがコメントアウト） |

## Subdirectories

なし。

## For AI Agents

### Working In This Directory

- Excel は **NetOfficeFw.Excel 1.9.10**（`NetOffice.ExcelApi`）。Interop COM 参照は使わない。
- 不要な Application / Workbook は `Dispose()`（`ComHelper.ReleaseCom` が IDisposable を Dispose）。
- 実行中の Excel は `Excel.Application.GetActiveInstance(false)`（`ExcelAppHelper`）。ユーザーの Excel を落とす `Quit` は呼ばない。
- `Application` を長時間保持しない。ブックが 0 件になったら `MainWindow.ReleaseExcelApp()` する（保持したままだと Excel が終了できない）。
- Excel 終了後に破棄済み RCW を触らない（`IsDisposed` / try-catch）。
- **キーボードフック** (`KeyboardHook` / `KeyHookManager`):
  - アプリ終了時に必ず `Close` / フック解除
  - ショートカット文字列形式は設定 YAML と一致させる（例: `"Alt + NumPad1"`）
- `ApplicationSetting.Location` は exe 基準の `ExcelTabSetting.yaml`。パス変更時は起動・保存の両方を確認
- Excel セル操作は `ExcelHelper`、お気に入りは `FavoriteHelper`、キャプチャは `CaptureHelper`、COM 解放は `ComHelper`。`Common` に戻さない。
- 例外は空 catch せず `AppLog.Warn` / `AppLog.Error` を使う。
- namespace が混在: `ApplicationSetting` と `Common` は `ExcelTab`、その他多くは `ExcelTab.CLASSES`

### Testing Requirements

- ROT: 複数ブック／アドイン（RelaxTools 等）開閉時にタブ一覧が正しいか
- フック: フォーカスが Excel / 他アプリのときでも意図したキーだけ反応するか
- 設定: Save 後の YAML 内容と再 Load

### Common Patterns

- Excel: `using Excel = Microsoft.Office.Interop.Excel;`
- 同期: `ROTManager.GetOpendLocalWorkbooks` に `[MethodImpl(MethodImplOptions.Synchronized)]`
- キー表現: カンマ区切りキー名 → `GetKeyToString` で `Ctrl + Alt + ...` 形式

## Dependencies

### Internal

- `ExcelTab.ITEM`（FavoriteExcel 等）— ApplicationSetting
- `ExcelTab.VIEW` / `MainWindow` — KeyHookManager から UI 操作
- `ExcelTab.EXTENTION` — 文字列配列 Trim 等

### External

- `Microsoft.Office.Interop.Excel`（COM）
- `YamlDotNet`
- `MouseKeyboardActivityMonitor` / Win32 user32 等
- `System.Windows` / Forms（キー列挙）

<!-- MANUAL: Any manually added notes below this line are preserved on regeneration -->
