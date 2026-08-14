<!-- Parent: ../AGENTS.md -->
<!-- Generated: 2026-08-10 | Updated: 2026-08-10 -->

# Properties

## Purpose

.NET / WPF プロジェクトのアセンブリメタデータ、埋め込みリソース、アプリケーションマニフェスト、デザイナー生成 Settings を格納する標準フォルダ。

## Key Files

| File | Description |
|------|-------------|
| `AssemblyInfo.cs` | アセンブリ名、バージョン、著作権等（例: 1.3.2.0, J.Handa） |
| `app.manifest` | 実行マニフェスト（UAC / 互換性）。csproj の `ApplicationManifest` から参照 |
| `Resources.resx` / `Resources.Designer.cs` | 埋め込みリソース（Designer は自動生成） |
| `Settings.settings` / `Settings.Designer.cs` | アプリケーション設定スケルトン（Designer は自動生成） |

## Subdirectories

なし。

## For AI Agents

### Working In This Directory

- `*.Designer.cs` は **手編集しない**（再生成で上書きされる）。
- バージョン番号を上げる場合は `AssemblyInfo.cs`（および必要なら表示用 `Common.GetVersion` 経路）を確認。
- `app.manifest` の変更は権限・DPI・互換性に影響する。安易に requireAdministrator 等を付けない。
- アプリの実ユーザー設定はここではなく **`ExcelTabSetting.yaml`**（YamlDotNet）が本命。

### Testing Requirements

- マニフェスト変更後は起動と Excel COM 連携を再確認
- バージョン表示（MainWindow メニュー）が期待どおりか

### Common Patterns

- Visual Studio の標準 Properties レイアウト

## Dependencies

### Internal

- csproj が Compile / EmbeddedResource / None として参照

### External

- .NET Framework ビルドシステム

<!-- MANUAL: Any manually added notes below this line are preserved on regeneration -->
