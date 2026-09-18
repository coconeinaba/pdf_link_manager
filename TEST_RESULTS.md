# テスト結果

実施日: 2026-08-21

## 検証環境

- Windows 64-bit
- Python 3.12.13
- 安全確認版: PyMuPDF 1.27.2.3 / MuPDF 1.27.2
- 比較対象: PyMuPDF 1.28.2 / MuPDF 1.28.2

## 結果

| 検査 | 結果 |
| --- | --- |
| Python全4ファイルの構文解析 | PASS |
| PyMuPDF 1.27.2.3の版検査 | PASS |
| PyMuPDF 1.28.2の起動拒否 | PASS |
| URLリンク1件の追加・再読込 | PASS |
| リンク追加前後のコンテンツストリームSHA-256 | 一致 |
| リンク追加前後の文字抽出 | `faeces`のまま一致 |
| リンク追加前後の画像差分 | 差分なし |
| Tkinter画面の生成・破棄スモークテスト | PASS |

自動テスト3件はすべて合格しました。

## 最新版での回帰再現

PyMuPDF 1.28.2でIssue #5054の最小PDFへ内容ストリーム再生成を伴う処理を行うと、抽出文字が次のように変化しました。

```text
faeces -> ffaeces
```

このため、現時点ではPyMuPDF 1.27.2.3固定を維持します。Issueの修正が正式リリースされた場合も、同じ回帰テストを通過してから固定版を更新してください。

## 実行コマンド

```powershell
python -m unittest discover -s tests -v
```
