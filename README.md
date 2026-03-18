# pptx-reviewer — PowerPoint レビュースキル for Claude Code

顧客向けPowerPoint資料（.pptx）を提出前に多角的にレビューし、改善ポイントをMarkdownレポートとして自動出力する **Claude Code スキル**です。

---

## 主な機能

| 観点 | 内容 |
|------|------|
| 📝 文章校正・表記ゆれ | 誤字脱字の検出、用語統一リストに基づく表記ゆれの指摘 |
| 🔗 論理的整合性 | 主語・述語のねじれ、因果関係の破綻、スライド間の矛盾の指摘 |
| 📊 構成・可読性 | 1スライドあたりの情報量（文字数）の偏りチェック、見出し階層の確認 |
| 👥 顧客向け表現 | 未説明の専門用語（略語）の指摘、です・ます調とだ・である調の混在チェック |

**対応図形**: テキストボックス / プレースホルダー / AutoShape（四角・丸・ダイヤなど）/ 表（テーブル）

---

## 必要環境

| ソフトウェア | バージョン |
|-------------|-----------|
| [Claude Code](https://github.com/anthropics/claude-code) | 最新版 |
| Python | 3.8 以上 |
| python-pptx | 0.6.21 以上（初回自動案内） |

> **Note**: `python-pptx` が未インストールの場合、スキル実行時にインストールを案内します。プロキシが必要な環境もサポートしています。

---

## インストール

### 1. リポジトリをクローン（またはダウンロード）

```bash
git clone https://github.com/<your-username>/pptx-reviewer.git
```

### 2. Claude Code のスキルディレクトリに配置

Claude Code のスキルは `~/.claude/skills/` 以下に置くことで自動認識されます。

```bash
# macOS / Linux
cp -r pptx-reviewer ~/.claude/skills/

# Windows (PowerShell)
Copy-Item -Recurse pptx-reviewer "$env:USERPROFILE\.claude\skills\"
```

### 3. 用語統一リストをカスタマイズ（任意）

`references/terminology.json` を編集して、自社の表記ルールを追加します（詳細は[用語リストのカスタマイズ](#用語統一リストのカスタマイズ)を参照）。

---

## 使い方

Claude Code のチャットで、ファイルパスとともにレビューを依頼するだけです。

### 全スライドをレビュー

```
proposal.pptx をレビューしてください。
```

```
C:\Users\username\Documents\proposal.pptx をすべてのページレビューしてください。
```

### 特定ページのみレビュー

ページ番号をカンマ区切りで指定します。

```
proposal.pptx の 2,5,6 ページをレビューしてください。
```

```
proposal.pptx の 3,4,5 ページをレビューしてください。特に顧客向け表現と論理的整合性を重点的に確認してください。
```

---

## 出力ファイル

レビュー結果は **PowerPointファイルと同じフォルダ** に保存されます。

### ファイル名の規則

```
YYYYMMDD_<元ファイル名>.md
```

| 元ファイル | 出力ファイル |
|-----------|-------------|
| `proposal.pptx` | `20260318_proposal.md` |
| `proposal_v2.pptx` | `20260318_proposal_v2.md` |

同名ファイルが存在する場合は末尾に連番が付与されます（`_01`, `_02`, ...）。

### レポート構成

```markdown
# PowerPoint レビューレポート

**ファイル**: proposal.pptx
**レビュー日時**: 2026-03-18
**総スライド数**: 10 枚
**レビュー対象ページ**: 2, 5, 6 ページ
**対象読者**: 顧客向け

## 総合サマリー
...

## スライド別指摘事項

### スライド 2：現状のシステム構成

#### 📝 文章校正・表記ゆれ
| 指摘内容 | 箇所 | 改善案 |
|...

#### 🔗 論理的整合性
...

## 全体的な改善提案
...
```

---

## 用語統一リストのカスタマイズ

`references/terminology.json` に自社の用語ルールを追加することで、スキルが自動的に表記ゆれを検出します。

### フォーマット

```json
{
  "version": "1.0",
  "description": "表記ゆれ・用語統一リスト",
  "last_updated": "2026-03-18",
  "terms": [
    {
      "correct": "サーバ",
      "variants": ["サーバー"],
      "category": "IT用語",
      "notes": "末尾長音符を省略する表記を正とする"
    },
    {
      "correct": "ユーザ",
      "variants": ["ユーザー"],
      "category": "IT用語",
      "notes": "末尾長音符を省略する表記を正とする"
    }
  ]
}
```

### フィールド説明

| フィールド | 必須 | 説明 |
|-----------|------|------|
| `correct` | ✅ | 正式表記（統一したい表記） |
| `variants` | ✅ | 誤表記・ゆれのリスト（複数指定可） |
| `category` | — | 用語のカテゴリ（任意、管理用） |
| `notes` | — | 注記・判断根拠（任意） |

### カスタマイズ例

```json
{
  "terms": [
    { "correct": "サーバ",    "variants": ["サーバー"],    "category": "IT用語" },
    { "correct": "ユーザ",    "variants": ["ユーザー"],    "category": "IT用語" },
    { "correct": "ネットワーク", "variants": ["ネットワーク"], "category": "IT用語" },
    { "correct": "インタフェース", "variants": ["インターフェース", "インターフェイス"], "category": "IT用語" },
    { "correct": "弊社",      "variants": ["当社", "自社"], "category": "ビジネス用語", "notes": "顧客向け文書では弊社を使う" }
  ]
}
```

---

## ファイル構成

```
pptx-reviewer/
├── SKILL.md                        # スキル本体（Claude への指示書）
├── README.md                       # このファイル
├── scripts/
│   ├── extract_pptx.py             # .pptx からテキスト・構造情報を抽出
│   ├── check_terminology.py        # 用語統一リストとの照合
│   └── create_dummy_pptx.py        # テスト用ダミーPPTX生成スクリプト
├── references/
│   └── terminology.json            # 用語統一リスト（カスタマイズして使用）
└── evals/
    └── evals.json                  # スキルのテストケース定義
```

### 各スクリプトの単体実行

スクリプトは独立して実行できます。デバッグや動作確認に使用してください。

```bash
# テキスト抽出（全スライド）
python scripts/extract_pptx.py path/to/file.pptx

# テキスト抽出（特定ページのみ）
python scripts/extract_pptx.py path/to/file.pptx --pages 2,5,6

# 用語チェック
python scripts/check_terminology.py extracted.json references/terminology.json

# テスト用ダミーPPTX生成
python scripts/create_dummy_pptx.py test.pptx
```

---

## 対応する図形の種類

`extract_pptx.py` が抽出する図形と `shape_kind` ラベルの対応表です。

| 図形の種類 | shape_kind | 説明 |
|-----------|-----------|------|
| タイトルプレースホルダー | `タイトル` | スライドタイトル |
| コンテンツプレースホルダー | `コンテンツ` | 箇条書きなどの本文エリア |
| テキストボックス | `テキストボックス` | 独立して配置されたテキストボックス |
| AutoShape（四角・丸・ダイヤなど） | `図形` | テキストを含む図形全般 |
| 表（テーブル） | `表` | セル内テキストを個別に抽出 |
| 画像 | —（スキップ） | テキストなしのため抽出対象外 |

---

## テスト用ダミーファイルの生成

動作確認用のPPTXファイルを生成できます。意図的な問題を含む8スライドのサンプルが作成されます。

```bash
python scripts/create_dummy_pptx.py sample.pptx
```

**含まれる問題点:**

| スライド | 問題の種類 |
|---------|-----------|
| スライド 2 | 表記ゆれ（サーバ/サーバー、ユーザ/ユーザー） |
| スライド 3 | 主語・述語のねじれ、因果関係の矛盾 |
| スライド 4 | 情報量過多（578文字） |
| スライド 5 | 専門用語の未説明（RBAC, APIM, IaC, ZTNA, SIEM） |
| スライド 6 | トーン不統一（です・ます ↔ だ・である） |
| スライド 8 | 四角・丸・ダイヤのAutoShape＋表内の表記ゆれ・専門用語 |

---

## 動作のしくみ

```
ユーザーの依頼
    │
    ▼
① ファイルパスとページ指定の確認
    │
    ▼
② Python / python-pptx の環境確認・インストール案内
    │
    ▼
③ terminology.json の読み込み
    │
    ▼
④ extract_pptx.py でテキスト・構造情報を抽出（JSON）
    │
    ▼
⑤ check_terminology.py で表記ゆれを自動検出
    │
    ▼
⑥ 4観点でレビュー分析（Claude による自然言語処理）
    │
    ▼
⑦ Markdownレポートを元のPPTXと同じフォルダに保存
```

---

## ライセンス

MIT License

---

## 貢献・フィードバック

Issue や Pull Request を歓迎します。
用語リストのサンプル追加、対応言語の拡張、レポートフォーマットの改善などを随時受け付けています。
