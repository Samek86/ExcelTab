<!-- Parent: ../AGENTS.md -->
<!-- Generated: 2026-08-10 | Updated: 2026-08-10 -->

# ITEM

## Purpose

シリアライズ可能なドメインモデルと列挙型。設定 YAML やお気に入り機能のデータ形状を定義する小さなフォルダ。

## Key Files

| File | Description |
|------|-------------|
| `FavoriteExcel.cs` | お気に入りエントリ。フルパス、シート名、行・列、セル指定フラグ（`IsCell`） |
| `WinkEnum.cs` | ウィンドウ／表示関連の列挙（Wink 系） |

## Subdirectories

なし。

## For AI Agents

### Working In This Directory

- `FavoriteExcel` のプロパティを変更すると **YAML の互換性**と `FavoriteWindow` / `FavoriteItemControl` / `ApplicationSetting.FavoriteCellDic` に影響する。
- ロジックは置かず、POCO／enum に留める方針が望ましい。
- namespace: `ExcelTab.ITEM`

### Testing Requirements

- お気に入り追加 → Save → 再起動 → 復元
- セルジャンプとシートのみジャンプ（`IsCell`）の分岐

### Common Patterns

- 単純な自動プロパティ
- `ApplicationSetting` から Dictionary で保持

## Dependencies

### Internal

- 参照される側: `ApplicationSetting`, VIEW のお気に入り UI

### External

- BCL のみ（`System.Drawing` 参照はファイル内 using 程度）

<!-- MANUAL: Any manually added notes below this line are preserved on regeneration -->
