# GoogleドキュメントからGoogleスプレッドシートへの階層的データ変換スクリプト

このGoogle Apps Scriptは、指定されたGoogleドキュメントの内容をマークダウン形式に変換し、Googleスプレッドシートに階層的に出力するためのものです。見出し、本文、および箇条書きの階層構造を維持しながら、スプレッドシートの各カラムに整理して格納します。

## 機能

- Googleドキュメントの見出しを階層ごとに判別し、適切なカラムに格納。
- 本文を「本文」カラムに、箇条書きを「箇条書き」カラムに行ごとに出力。
- GoogleドキュメントとGoogleスプレッドシートのURL指定が可能。

## 前提条件

- GoogleアカウントおよびGoogle Apps Scriptへのアクセス権。
- Googleドキュメント（見出しと箇条書きの形式を含む）が必要。
- データの出力先となるGoogleスプレッドシートが必要。

## インストール

1. [Google Apps Script](https://script.google.com/)にアクセスし、新規プロジェクトを作成。
2. コードエディタにスクリプトを貼り付け、`docUrl`と`sheetUrl`をそれぞれのGoogleドキュメントとGoogleスプレッドシートのURLに置き換えます。
3. プロジェクトを保存し、実行時に必要なアクセス権限を許可します。

## 使い方

1. `docUrl`と`sheetUrl`を、それぞれGoogleドキュメントとGoogleスプレッドシートのURLに設定します。
2. `importGoogleDocToSheetByUrl` 関数を実行すると、Googleドキュメントの内容が解析され、Googleスプレッドシートに出力されます。
3. スクリプトは、見出し、本文、および箇条書きごとに行を分けて階層的にスプレッドシートに格納します。

## 出力形式の例

Googleドキュメントが次のような構成の場合:

```markdown
markdown
Copy code
# 見出し1
## 見出し1-1
DETAIL: 本文1
- 箇条書き1-1
- 箇条書き1-2
## 見出し1-2
DETAIL: 本文2
- 箇条書き2-1
DETAIL: 本文3
- 箇条書き3-1
- 箇条書き3-2

```

スプレッドシートの出力:

| 見出し1 | 見出し2 | 本文 | 箇条書き |
| --- | --- | --- | --- |
| 見出し1 | 見出し1-1 | 本文1 | 箇条書き1-1 |
| 見出し1 | 見出し1-1 | 本文1 | 箇条書き1-2 |
| 見出し1 | 見出し1-2 | 本文2 | 箇条書き2-1 |
| 見出し1 | 見出し1-2 | 本文3 | 箇条書き3-1 |
| 見出し1 | 見出し1-2 | 本文3 | 箇条書き3-2 |

## 関数の詳細

### `importGoogleDocToSheetByUrl()`

- **概要**: Googleドキュメントから内容を取得し、マークダウン形式に変換後、スプレッドシートに出力します。
- **パラメータ**: なし（URLは関数内で設定）
- **注意**: 実行前にGoogleドキュメントとGoogleスプレッドシートのURLを設定してください。

### `convertMarkdownToSheet(markdownText, sheetId)`

- **概要**: マークダウン形式のテキストを解析し、見出し、本文、箇条書きごとに行を分けてスプレッドシートに出力します。
- **パラメータ**:
    - `markdownText`: マークダウン形式に変換されたテキスト
    - `sheetId`: スプレッドシートのID
- **注意**: スプレッドシートの既存のデータはクリアされます。

### `extractIdFromUrl(url)`

- **概要**: GoogleドキュメントやGoogleスプレッドシートのURLからIDを抽出します。
- **パラメータ**:
    - `url`: GoogleドキュメントまたはスプレッドシートのURL
- **返り値**: 抽出されたID

## 注意点

- スクリプト実行時、スプレッドシートの内容はクリアされます。データがある場合は、事前にバックアップを取ってください。
- Googleドキュメントとスプレッドシートの両方でアクセス権限が必要です。

## ライセンス

このスクリプトは自由に利用および改変できます。使用は自己責任で行ってください。
