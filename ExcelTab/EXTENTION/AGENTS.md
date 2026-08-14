<!-- Parent: ../AGENTS.md -->
<!-- Generated: 2026-08-10 | Updated: 2026-08-10 -->

# EXTENTION

## Purpose

C# 拡張メソッド置き場。フォルダ名は歴史的に **EXTENTION** のまま。型は `ExcelTab.Extension.StringExtension`。

## Key Files

| File | Description |
|------|-------------|
| `StringExtension.cs` | `string[]` 向け `Trim()` 拡張（全要素を Trim） |

## Subdirectories

なし。

## For AI Agents

### Working In This Directory

- 新規拡張は関連する型ごとにファイル分割を推奨。
- namespace: `ExcelTab.Extension`
- 利用箇所の例: `KeyHookManager` が `using ExcelTab.EXTENTION` でキー配列を Trim

### Testing Requirements

- 変更時は呼び出し元（キー文字列パース等）の回帰を手動確認

### Common Patterns

- `public static class` + `this` 引数

## Dependencies

### Internal

- 呼び出し元: 主に `CLASSES/KeyHookManager.cs`

### External

- BCL（LINQ）

<!-- MANUAL: Any manually added notes below this line are preserved on regeneration -->
