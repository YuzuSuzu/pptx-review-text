# pptx-reviewer — PowerPoint レビューエージェント（Codex CLI 用）

ユーザーが PowerPoint ファイル（.pptx）のレビュー・チェック・添削を依頼したとき、
または「pptx を見て」「スライドを確認して」「資料をレビューして」という表現が出たときは、
以下の手順に従ってレビューを実行してください。

## 対象

- 主に日本語で書かれた PowerPoint 資料（.pptx）
- システムエンジニアが作成した顧客向け資料を想定（英単語・英語の関数名・技術用語が混在する場合も許容）
- 対象読者：顧客（IT 専門家でない可能性がある）

---

## スクリプトの場所

このスキルは以下の標準パスに配置されていることを前提とします。

| OS | スキルフォルダのパス |
|----|--------------------|
| Windows | `%USERPROFILE%\.codex\skills\pptx-reviewer` |
| macOS / Linux | `~/.codex/skills/pptx-reviewer` |

スクリプトは以下のパスにあります（パス修正不要・自動検出）：

| スクリプト | 用途 |
|-----------|------|
| `~/.codex/skills/pptx-reviewer/scripts/extract_pptx.py` | .pptx からテキスト・構造情報を抽出 |
| `~/.codex/skills/pptx-reviewer/scripts/check_terminology.py` | 用語統一リストとの照合・表記ゆれ検出（terminology.json を自動検出） |
| `~/.codex/skills/pptx-reviewer/references/terminology.json` | 用語統一リスト |

> `check_terminology.py` は隣接する `references/terminology.json` を自動的に参照するため、パスの指定は不要です。

---

## Step 1: ファイルとレビュー対象ページの確認

ユーザーから PowerPoint ファイルのパスを受け取る。パスが指定されていない場合は確認する。

**ページ指定について：**
- ページ（スライド番号）が指定されている場合は、そのページのみをレビュー対象とする
- ページ指定はカンマ区切りで複数指定可能（例：`1,3,5,7`）
- 指定なしの場合は全スライドをレビューする

ページ指定の受け取り方の例：
- `proposal.pptx の 2,4,7 ページだけレビューして`
- 指定なし → 全スライドが対象

ページ指定がある場合は、Step 4 のテキスト抽出時に `--pages` オプションで渡す。

> **[進捗通知]** Step 1 完了後、ユーザーに以下を伝える：
> 「ファイルを確認しました。`<ファイル名>`（対象: <ページ指定 or 全スライド>）のレビューを開始します。」

---

## Step 2: Python 環境の確認と python-pptx のインストール

```bash
python --version
python -c "import pptx; print('ok')"
```

エラーが出た場合（未インストール）：

```bash
pip install python-pptx
```

インストール後、再度 `import pptx` で確認する。

> **[進捗通知]** Step 2 完了後、ユーザーに以下を伝える：
> 「Python環境を確認しました。テキスト抽出を準備しています...」

---

## Step 3: 用語統一リストの読み込み

`~/.codex/skills/pptx-reviewer/references/terminology.json` を読み込み、正式表記と誤表記（ゆれ）のリストを把握する。
このファイルは `check_terminology.py` が自動的に参照するため、コマンドへの指定は不要。

ファイルの構造：
```json
{
  "terms": [
    {
      "correct": "サーバ",
      "variants": ["サーバー"],
      "category": "IT用語",
      "notes": "末尾長音符を省略する表記を正とする"
    }
  ]
}
```

> **[進捗通知]** Step 3 完了後、ユーザーに以下を伝える：
> 「用語リストを読み込みました（N語）。テキスト抽出・用語チェックを実行します...」
> ※ 用語リストがない場合：「用語リストが見つかりません。一般的な表記ゆれの検出のみで続行します。」

---

## Step 4: テキスト抽出 → 用語チェックの一括実行

`extract_pptx.py` の出力を直接 `check_terminology.py` にパイプで渡す。**中間ファイルは作成しない。**

**bash / PowerShell（macOS・Linux・Windows PowerShell）：**

全スライドを対象にする場合：
```bash
python ~/.codex/skills/pptx-reviewer/scripts/extract_pptx.py <path-to-pptx> | python ~/.codex/skills/pptx-reviewer/scripts/check_terminology.py -
```

特定ページのみ対象にする場合（カンマ区切りでページ番号を指定）：
```bash
python ~/.codex/skills/pptx-reviewer/scripts/extract_pptx.py <path-to-pptx> --pages 1,3,5 | python ~/.codex/skills/pptx-reviewer/scripts/check_terminology.py -
```

**Windows コマンドプロンプト（CMD）の場合：**

```cmd
python %USERPROFILE%\.codex\skills\pptx-reviewer\scripts\extract_pptx.py <path-to-pptx> | python %USERPROFILE%\.codex\skills\pptx-reviewer\scripts\check_terminology.py -
```

> **注意（Windows）**: エンコーディングが乱れる場合は先頭に `set PYTHONUTF8=1` を付けてから実行する。ただし各スクリプトは起動時に自動でUTF-8設定を行うため、通常は不要。

出力される JSON の主なフィールド（extract_pptx.py より）：
- `total_slides`: ファイル全体のスライド枚数
- `reviewed_slides`: 実際にレビュー対象としたスライド番号のリスト
- `slides[].title`: スライドタイトル
- `slides[].total_chars`: そのスライドの総文字数
- `slides[].shapes[].shape_kind`: 図形の種別（タイトル／コンテンツ／テキストボックス／図形／表／グラフ／SmartArt／代替テキスト）
- `slides[].shapes[].paragraphs`: 段落テキストと見出しレベル（level）
- `slides[].notes`: ノートペイン（スピーカーノート）のテキスト。ノートがない場合は `null`

パイプが使えない環境では一時ファイルを経由してもよい。その場合はレビュー完了後に必ず削除すること：
```bash
python ~/.codex/skills/pptx-reviewer/scripts/extract_pptx.py <path-to-pptx> > extract_out.json
python ~/.codex/skills/pptx-reviewer/scripts/check_terminology.py extract_out.json
del extract_out.json
```

> **[進捗通知]** スクリプトの進捗は stderr に `[1/2]` `[2/2]` プレフィクスで自動出力される。
> スクリプト完了後、ユーザーに以下を伝える：
> 「テキスト抽出・用語チェックが完了しました。レビュー分析を開始します...（この処理は少し時間がかかります）」

---

## Step 5: 用語チェック結果の確認

前の手順で `check_terminology.py` が出力した JSON を確認する（stdout に出力されるため変数やメモリ上で扱う）。

出力 JSON には、スライドごとに「どの用語が、どのスライドの、どのコンテキストで誤表記されていたか」が含まれる。
この結果を Step 6 の観点1（文章校正・表記ゆれ）に組み込む。

> **[進捗通知]** Step 5 完了後、ユーザーに以下を伝える：
> 「レビュー分析が完了しました。レポートを生成・保存します...」

---

## Step 6: レビュー分析

抽出したテキスト全体と Step 5 の用語チェック結果を組み合わせて、以下の4つの観点でレビューを行う。

### 観点1：文章校正・表記ゆれ

- **用語統一リストによる検出**：check_terminology.py 結果を使い、正式表記でない用語を漏れなく報告する
- **リスト外の表記ゆれ**：スライド内で同じ概念に複数の表記が混在している場合も指摘する
- **誤字・脱字の検出**：明らかな入力ミスや不自然な文字の欠落
- **文法エラー**：助詞の誤り、読点の欠落、文として成立していない箇所
- `slides[].notes` にテキストがある場合はノートペインも対象とする（箇所は「ノートペイン」と明記）
- 英単語・英語の関数名・技術用語は**指摘対象外**

### 観点2：論理的整合性

- **スライド間の矛盾**：前後のスライドで事実が食い違っていないか
- **主語・述語のねじれ**：文の主語と述語が意味的に対応しているか
- **因果関係の破綻**：接続詞で繋がれた文が実際に因果関係を成していないケース

### 観点3：構成と可読性

- **情報量の偏り**：`total_chars` を参照し、300文字超はやや多め、600文字超は要注意として指摘
- **見出し階層の適切さ**：段落の `level` 情報から、階層が飛んでいないか確認
- **1スライド1メッセージ原則**：複数のテーマが混在しているスライドはスライド分割を提案

### 観点4：顧客向け表現の調整

- **過度な専門用語の指摘**：技術的な略語（例：APIM、RBAC、IaC など）を説明なしに使っている場合に指摘
- **トーン＆マナーの統一**：です・ます調とだ・である調の混在を指摘
- **曖昧な表現**：「など」「等」「場合によっては」が多用されている箇所

---

## Step 7: レポートの出力

レビュー結果を Markdown ファイルとして保存する。

### 出力ファイル名と保存先

- **保存先**: PowerPoint ファイルと**同じフォルダ**
- **ファイル名**: `YYYYMMDD_<元ファイル名（拡張子なし）>.md`
- **重複時**: 末尾に2桁の連番を付与（`_01`, `_02`, ...）

ファイル名を組み立てる手順（Python で実行）：

```python
import os
from datetime import date

pptx_path = "<PowerPointのパス>"
folder = os.path.dirname(os.path.abspath(pptx_path))
stem = os.path.splitext(os.path.basename(pptx_path))[0]
today = date.today().strftime("%Y%m%d")
base_name = f"{today}_{stem}.md"
out_path = os.path.join(folder, base_name)

if os.path.exists(out_path):
    i = 1
    while True:
        candidate = os.path.join(folder, f"{today}_{stem}_{i:02d}.md")
        if not os.path.exists(candidate):
            out_path = candidate
            break
        i += 1
```

### レポートの構成

**このフォーマットを毎回厳守すること。セクションの順序・表の列構成を変えてはならない。**

```markdown
# PowerPoint レビューレポート

**ファイル**: <ファイル名>
**レビュー日時**: <YYYY-MM-DD>
**総スライド数**: <N> 枚（ファイル全体）
**レビュー対象ページ**: 全スライド  ← ページ指定がある場合は「1, 3, 5 ページ」のように記載
**対象読者**: 顧客向け

---

## 総合サマリー

<全体を通じた主な課題と優先度が高い改善ポイントを3〜5行で要約する。
表記ゆれの件数・論理問題の有無・専門用語の多用・文体の乱れなど主要な問題を端的にまとめる>

---

## カテゴリ別 指摘件数

| カテゴリ | 件数 |
|---------|------|
| 📝 表記ゆれ・文章校正 | N件 |
| 🔗 論理的整合性 | N件 |
| 📊 構成・可読性 | N件 |
| 👥 専門用語・顧客表現 | N件 |
| **合計** | **N件** |

---

## スライド別 指摘事項

### スライド <N>：<タイトル>

| カテゴリ | 箇所 | 指摘内容 | 改善案 |
|---------|------|---------|--------|
| 📝 表記ゆれ | テキストボックス 2行目 | 「サーバー」→「サーバ」（用語リスト） | 「サーバ」に統一することを推奨 |
| 🔗 論理的整合性 | テキストボックス 1行目 | 主語「課題は」に対して述語「改善します」がねじれている | 「課題はパフォーマンスの低下です」などに修正することを推奨 |
| 📊 構成・可読性 | テキストボックス 全体 | 文字数約578文字。情報量が過多 | 2スライドへ分割することを推奨 |
| 👥 専門用語 | テキストボックス 項目1 | 「RBAC」が説明なしに使用されている | 「ロールベースアクセス制御（RBAC）」と初出時に説明を追加することを推奨 |

（指摘がないスライドはセクションごと省略する）

---

## 全体的な改善提案

1. <改善提案1（表記統一など最優先のもの）>
2. <改善提案2>
3. <改善提案3>
```

### 出力ルール

- **セクション順序を厳守**: 総合サマリー → カテゴリ別件数表 → スライド別指摘事項 → 全体的な改善提案
- **スライド別は必ず1つの統合テーブルで表現する**: カテゴリごとにセクションを分けない。列は「カテゴリ｜箇所｜指摘内容｜改善案」の4列固定
- **カテゴリ列のアイコン**: 📝 表記ゆれ・文章校正 / 🔗 論理的整合性 / 📊 構成・可読性 / 👥 専門用語・顧客表現 の4種のみ使用
- **テキストの引用・転載は禁止**: レポートには指摘内容と箇所の説明のみ記載する
- **箇所の示し方**: shape_kind（タイトル／テキストボックス／図形／表／グラフ／ノートペイン）と位置（「〇行目」「〇列目」など）で示す
- **指摘がないスライドは省略**: スライド別セクションに含めない
- **改善案は推奨トーン**: 「〜することを推奨」で統一。「禁止」「必須」などの断定表現は使わない
- **カテゴリ別件数は正確に集計**してから記載する

---

## Step 8: 完了後の案内

レポートファイルの保存先パスをユーザーに伝える：

> レビューが完了しました。結果は `<出力パス>` に保存しました。
> 主な指摘事項：<総合サマリーの1行要約>
