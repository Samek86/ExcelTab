<!-- Parent: ../AGENTS.md -->
<!-- Generated: 2026-08-10 | Updated: 2026-08-10 -->

# STYLE

## Purpose

アプリ全体およびコントロール向けの WPF `ResourceDictionary`（XAML スタイル定義）。ダーク基調の共通見た目を定義する。

## Key Files

| File | Description |
|------|-------------|
| `CommonStyle.xaml` | 共通スタイル（**App.xaml でマージ済み**） |
| `ContextMenuStyle.xaml` | コンテキストメニュー（マージ済み） |
| `ButtonStyle.xaml` | ボタン（マージ済み） |
| `ScrollStyle.xaml` | スクロールバー（マージ済み） |
| `RadioStyle.xaml` | ラジオボタン（マージ済み） |
| `TextboxStyle.xaml` | テキストボックス（マージ済み） |
| `TextblockStyle.xaml` | テキストブロック（マージ済み） |
| `TabStyle.xaml` | タブ用（App.xaml ではコメントアウト） |
| `LabelStyle.xaml` | ラベル（コメントアウト） |
| `ComboboxStyle.xaml` | コンボボックス（コメントアウト） |
| `MenuStyle.xaml` | メニュー（コメントアウト） |
| `ProgressBarStyle.xaml` | プログレスバー（コメントアウト） |
| `StepItemStyles.xaml` | ステップ項目用 |

## Subdirectories

なし。

## For AI Agents

### Working In This Directory

- **グローバルに効かせる**スタイルは `App.xaml` の `MergedDictionaries` に `Source="/STYLE/....xaml"` を追加する。
- ファイルだけ置いても未マージならアプリ全体には適用されない。
- キー名（`x:Key`）の衝突に注意。既存コントロールの見た目を壊さないよう、変更後は MainWindow・ダイアログを目視確認。
- ダーク UI の色は `Common` のブラシ定義とも連携している。色変更時は両方を確認。

### Testing Requirements

- ビルド後にボタン、メニュー、スクロール、テキスト入力の見た目とホバー／押下状態を確認
- 新規マージ時は起動エラー（ResourceDictionary パス不正）がないか確認

### Common Patterns

- 各ファイルは単一の `ResourceDictionary`
- パス表記: `/STYLE/{Name}.xaml`（アセンブリルート相対）

## Dependencies

### Internal

- `App.xaml` がマージのエントリポイント
- 各 VIEW / MainWindow のコントロールテンプレートが参照

### External

- WPF PresentationFramework

<!-- MANUAL: Any manually added notes below this line are preserved on regeneration -->
