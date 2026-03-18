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

## はじめてのセットアップ（Claude Code CLI 導入手順）

このスキルは **Claude Code CLI**（Anthropic 社製の AI コーディングアシスタント）と一緒に使います。
ここでは **Windows でゼロから使えるようにする手順**を初めての方向けに丁寧に説明します。

---

### STEP 1: Node.js のインストール確認

Claude Code は Node.js で動作します。まずコマンドプロンプトを開いて確認します。

> **コマンドプロンプトの開き方**: スタートメニューで「cmd」と検索 → 「コマンドプロンプト」を開く

```cmd
node --version
```

バージョン番号（例: `v20.11.0`）が表示されれば OK です。
**表示されない場合** → [Node.js 公式サイト](https://nodejs.org/ja/) から **「LTS 版（推奨版）」** をダウンロードしてインストールしてください。インストール中の設定はすべてデフォルト（Next ボタン連打）で問題ありません。

---

### STEP 2: Python のインストール確認

スライドのテキスト抽出に Python が必要です。

```cmd
python --version
```

`Python 3.x.x` と表示されれば OK です。
**表示されない場合** → [Python 公式サイト](https://www.python.org/downloads/) からインストールしてください。

> **重要**: インストール画面の最初に表示される **「Add Python to PATH」のチェックボックスを必ずオンにしてください。** これを忘れると Python が認識されません。

---

### STEP 3: Claude Code のインストール

Node.js が使えることを確認したら、以下のコマンドを実行します。

```cmd
npm install -g @anthropic-ai/claude-code
```

インストールが完了したら確認します。

```cmd
claude --version
```

バージョン番号が表示されれば成功です。

---

### STEP 4: Anthropic API キーの取得と設定

Claude Code を動かすには Anthropic の API キーが必要です。

#### 4-1. API キーを取得する

1. [Anthropic Console](https://console.anthropic.com/) にアクセスしてアカウントを作成・ログイン
2. 左メニューの **「API Keys」** をクリック
3. **「Create Key」** ボタンで任意の名前を入力して作成
4. 表示された `sk-ant-` で始まる文字列をコピーしてメモ帳などに保存

> **注意**: キーはこの画面を閉じると二度と表示されません。必ずすぐにコピーして保存してください。

#### 4-2. API キーを環境変数に設定する

コマンドプロンプトで以下を実行します（`ここにキーを貼り付ける` の部分を実際のキーに置き換えてください）。

```cmd
setx ANTHROPIC_API_KEY "sk-ant-ここにキーを貼り付ける"
```

設定後は **コマンドプロンプトを一度閉じて開き直します**（設定を反映させるため）。

正しく設定されたか確認します。

```cmd
echo %ANTHROPIC_API_KEY%
```

`sk-ant-...` が表示されれば設定完了です。

---

### STEP 5: このスキルのインストール

#### 5-1. スキルをダウンロードする

**git を使う場合（推奨）：**

```cmd
git clone https://github.com/YuzuSuzu/pptx-review-text.git
```

**git がない場合：**
このページ上部の緑色の **「Code」ボタン → 「Download ZIP」** でダウンロードして解凍してください。

#### 5-2. スキルフォルダに配置する

Claude Code はスキルを以下のフォルダから自動認識します。

```
C:\Users\（あなたのユーザー名）\.claude\skills\
```

以下のコマンドでフォルダを作成してコピーします。

```cmd
:: スキル用フォルダを作成（初回のみ）
mkdir "%USERPROFILE%\.claude\skills"

:: クローン（またはダウンロード）したフォルダをコピー
xcopy /E /I pptx-review-text "%USERPROFILE%\.claude\skills\pptx-reviewer"
```

コピー後の確認：

```
C:\Users\（ユーザー名）\
└── .claude\
    └── skills\
        └── pptx-reviewer\        ← このフォルダが Claude Code に認識される
            ├── SKILL.md
            ├── scripts\
            └── references\
```

#### 5-3. python-pptx ライブラリをインストールする

```cmd
pip install python-pptx
```

正しくインストールされたか確認します。

```cmd
python -c "import pptx; print('OK')"
```

`OK` と表示されれば成功です。

---

### STEP 6: 実際に使ってみる

#### 6-1. Claude Code を起動する

レビューしたい PPTX ファイルが入っているフォルダに移動してから起動します。

```cmd
cd C:\Users\（ユーザー名）\Documents
claude
```

チャット画面が表示されれば起動成功です。

#### 6-2. レビューを依頼する

チャット画面に話しかけるだけで動作します。

**全スライドをレビューする場合：**

```
proposal.pptx をレビューしてください。
```

**特定のページだけレビューする場合：**

```
proposal.pptx の 2,5,6 ページをレビューしてください。
```

**重点項目を指定する場合：**

```
proposal.pptx をレビューしてください。特に専門用語の未説明と表記ゆれを重点的に確認してください。
```

Claude Code がスキルを自動認識してレビューを実行し、PPTX と同じフォルダに Markdown レポートを保存します。

---

### よくあるトラブルと対処法

| 症状 | 原因と対処 |
|------|-----------|
| `claude: コマンドが見つかりません` | `npm install -g @anthropic-ai/claude-code` を再実行。それでも出る場合はコマンドプロンプトを閉じて開き直す |
| `API key not found` | STEP 4-2 を確認。コマンドプロンプトを閉じて開き直す |
| `python: コマンドが見つかりません` | Python のインストールを確認。インストール時に「Add Python to PATH」にチェックが入っていたか確認し、必要に応じて再インストール |
| `ModuleNotFoundError: No module named 'pptx'` | `pip install python-pptx` を実行 |
| スキルが認識されない（レビューが動かない） | `%USERPROFILE%\.claude\skills\pptx-reviewer\SKILL.md` が存在するか確認 |
| プロキシ環境で pip が失敗する | `pip install python-pptx --proxy http://プロキシアドレス:ポート番号` を使う |

---

## インストール（上級者向け・要約）

### 1. リポジトリをクローン（またはダウンロード）

```bash
git clone https://github.com/YuzuSuzu/pptx-review-text.git
```

### 2. Claude Code のスキルディレクトリに配置

Claude Code のスキルは `~/.claude/skills/` 以下に置くことで自動認識されます。

```bash
# macOS / Linux
cp -r pptx-review-text ~/.claude/skills/pptx-reviewer

# Windows (PowerShell)
Copy-Item -Recurse pptx-review-text "$env:USERPROFILE\.claude\skills\pptx-reviewer"
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
