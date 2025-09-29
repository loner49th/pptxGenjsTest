# txt2ppt-cli

Markdownで書いたスライド原稿を[PptxGenJS](https://gitbrent.github.io/PptxGenJS/)でPowerPoint（.pptx）に変換するTypeScript製CLIです。slides.mdのようなテキストから版面を自動レイアウトし、出力先のプレゼン資料を手早く生成できます。

## 特徴
- 見出し（#）をスライドタイトル、##をサブタイトルとして解釈
- 段落・箇条書き（入れ子は空白2文字=1段）・番号付きリスト・コードブロック・表・画像をサポート
- 画像は![alt](path#cover|contain)でトリミング方式を指定可能
- >note:でスピーカーノート、>bg:でスライド背景画像を指定
- 版面の安全マージンを自動計算し、縦方向に収まらない場合は続きのスライドを分割生成
- 日本語向けにデフォルトフォントを指定しつつ、プレゼンのメタ情報（タイトル/著者/会社）を埋め込むオプション付き

## 必要条件
- Node.js 18以降（動作確認: Node.js 20系）
- npm 8以降

## セットアップ
`bash
npm install
`
依存パッケージ（TypeScript, ts-node, Jestなど）がインストールされます。

## 使い方
### コマンド
開発中はts-node経由で直接実行します。
`bash
npx ts-node txt2ppt.ts --in slides.md --out deck.sample.pptx
`
package.jsonには同じコマンドを実行するスクリプトが用意されています。
`bash
npm run txt2ppt
`

### オプション
| オプション | 説明 |
| --- | --- |
| --in <path> | 入力Markdownファイルへのパス（必須） |
| --out <path> | 出力する.pptxファイルパス（必須） |
| --layout <name> | スライドレイアウト。LAYOUT_16x9（既定）/LAYOUT_4x3/LAYOUT_WIDE/LAYOUT_16x10 |
| --title <text> | プレゼン全体のタイトルメタ情報 |
| --author <text> | 著者メタ情報 |
| --company <text> | 会社名メタ情報 |
| --bg <path> | 全スライド共通の背景画像パス |

### Markdown記法サポート
- **スライド区切り**: # で新しいスライドを開始。連続する#が出現すると自動的に前スライドを確定。
- **サブタイトル**: ## Subtitle
- **段落**: 空行で区切った通常テキスト。
- **箇条書き**: - item / * item。インデント2スペースごとにレベルを下げられます。番号付きリスト（1. item）にも対応。
- **コードブロック**: `` `lang ... ` ``。langはPptxGenJSのハイライト指定に使用。
- **表**: | A | B | 形式の行を連続させるとテーブルとして描画。
- **画像**: ![代替テキスト](path/to/image.png#cover) のように指定。#coverまたは#containを付けるとサイズ調整を制御。
- **ノート**: >note: ここに話者メモ。最初のスライドにまとめて書き出します。
- **背景**: >bg: path/to/background.png。スライド単位で背景を上書き。

## サンプルワークフロー
1. slides.mdを編集し、上記記法でスライド構成を記述
2. pm run txt2pptを実行しdeck.sample.pptxを生成
3. PowerPointで開き、必要に応じて最終調整

## 開発メモ
- 型定義や描画ロジックは	xt2ppt.tsにまとまっています。
- レイアウト計算は安全マージンと行間を考慮し、はみ出したブロックは自動で次スライドへ分割されます。
- テストフレームワークとしてJestを導入済みです。必要に応じて
pm testでユニットテストを追加・実行してください。
