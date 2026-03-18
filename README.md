# pptx-reviewer — PowerPoint レビュースキル

顧客向けPowerPoint資料（.pptx）を提出前に多角的にレビューし、改善ポイントをMarkdownレポートとして自動出力するスキルです。

**Claude Code（Anthropic）と Codex CLI（OpenAI）の両方に対応しています。**

| ツール | スキル定義ファイル | 使用 AI |
|--------|-----------------|---------|
| Claude Code CLI（Anthropic） | `SKILL.md` | Claude |
| Codex CLI（OpenAI） | `AGENTS.md` | GPT / o3 系 |

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

### Claude Code（Anthropic）で使う場合

| ソフトウェア | バージョン |
|-------------|-----------|
| [Claude Code](https://github.com/anthropics/claude-code) | 最新版 |
| Python | 3.8 以上 |
| python-pptx | 0.6.21 以上 |

### Codex CLI（OpenAI）で使う場合

| ソフトウェア | バージョン |
|-------------|-----------|
| [Codex CLI](https://github.com/openai/codex) | 最新版 |
| Python | 3.8 以上 |
| python-pptx | 0.6.21 以上 |

> **Note**: `python-pptx` が未インストールの場合はスキル実行前に `pip install python-pptx` を実行してください。

---

## このスキルを Claude Code CLI に追加する方法（導入済みの方向け）

> Claude Code CLI がまだ入っていない方は、次の「[はじめてのセットアップ](#はじめてのセットアップclaude-code-cli-導入手順)」を先にご覧ください。

---

### STEP 1: このリポジトリを取得する

ダウンロードしたいフォルダに移動してから、以下のどちらかの方法で取得します。

**方法 A：git でクローンする（推奨）**

```cmd
git clone https://github.com/YuzuSuzu/pptx-review-text.git
```

実行すると、現在のフォルダの中に `pptx-review-text` というフォルダが作られます。

**方法 B：ZIP でダウンロードする（git がない場合）**

1. このページ上部の緑色の **「Code」** ボタンをクリック
2. **「Download ZIP」** をクリック
3. ダウンロードした ZIP を右クリック →「すべて展開」で解凍
4. 解凍してできたフォルダ名を `pptx-reviewer` に変更する

---

### STEP 2: スキルフォルダに配置する

Claude Code は **ホームディレクトリの `.claude/skills/` フォルダ**の中を自動でスキルとして認識します。
ここに `pptx-reviewer` という名前のフォルダを置くだけで使えるようになります。

#### スキルフォルダの場所（OS 別）

| OS | スキルフォルダのパス |
|----|--------------------|
| Windows | `C:\Users\（ユーザー名）\.claude\skills\` |
| macOS / Linux | `~/.claude/skills/` |

> **`.claude` フォルダが見えない場合（Windows）**：エクスプローラーの「表示」→「隠しファイル」にチェックを入れると表示されます。

#### Windows の場合（コマンドプロンプト）

```cmd
:: スキルフォルダがなければ作成（初回のみ）
mkdir "%USERPROFILE%\.claude\skills"

:: pptx-reviewer をコピー（方法Aでクローンした場合）
xcopy /E /I pptx-review-text "%USERPROFILE%\.claude\skills\pptx-reviewer"
```

方法 B（ZIP 解凍・リネーム済み）の場合は、エクスプローラーで `pptx-reviewer` フォルダを以下の場所に移動するだけでも OK です。

```
C:\Users\（ユーザー名）\.claude\skills\
```

#### macOS / Linux の場合（ターミナル）

```bash
mkdir -p ~/.claude/skills
cp -r pptx-review-text ~/.claude/skills/pptx-reviewer
```

---

### STEP 3: 配置後のフォルダ構成を確認する

配置が正しければ、以下のような構成になっているはずです。

```
C:\Users\（ユーザー名）\           ←「ホームディレクトリ」と呼ぶ場所
└── .claude\
    └── skills\
        └── pptx-reviewer\               ← このフォルダ名が重要
            ├── SKILL.md                 ← Claude Code はここを読んでスキルを認識する ★必須
            ├── README.md
            ├── scripts\
            │   ├── extract_pptx.py
            │   ├── check_terminology.py
            │   └── create_dummy_pptx.py
            ├── references\
            │   └── terminology.json
            └── evals\
                └── evals.json
```

> **チェックポイント（Claude Code）**：`SKILL.md` が `pptx-reviewer\` の直下にあることが必須です。ここにないと Claude Code がスキルを認識できません。

---

### Codex CLI の場合の配置方法

Codex CLI は **`~/.codex/`（Windowsでは `C:\Users\（ユーザー名）\.codex\`）** をグローバル設定フォルダとして使います。
Claude Code の `~/.claude/skills/` と同じ考え方で、ここに `AGENTS.md` を置くと**どのフォルダから `codex` を起動しても自動で読み込まれます**。

#### 推奨の配置構成

```
C:\Users\（ユーザー名）\
└── .codex\
    ├── AGENTS.md                    ← Codex CLI がここを自動で読み込む ★
    └── pptx-reviewer\               ← スクリプト・用語リストの置き場
        ├── SKILL.md
        ├── scripts\
        └── references\
```

#### 配置コマンド（Windows）

```cmd
:: .codex フォルダを作成（初回のみ）
mkdir "%USERPROFILE%\.codex"

:: リポジトリを .codex\pptx-reviewer として配置
xcopy /E /I pptx-review-text "%USERPROFILE%\.codex\pptx-reviewer"

:: AGENTS.md を .codex 直下にコピー
copy "%USERPROFILE%\.codex\pptx-reviewer\AGENTS.md" "%USERPROFILE%\.codex\AGENTS.md"
```

#### AGENTS.md 内のパスを更新する

コピーした `%USERPROFILE%\.codex\AGENTS.md` をメモ帳で開き、`SKILL_DIR` を一括置換します。

```
変更前： SKILL_DIR
変更後： C:\Users\（ユーザー名）\.codex\pptx-reviewer
```

> **すでに `.codex\AGENTS.md` がある場合** は上書きせず、ファイルの末尾に内容を**追記**してください。

#### macOS / Linux の場合

```bash
mkdir -p ~/.codex
cp -r pptx-review-text ~/.codex/pptx-reviewer
cat ~/.codex/pptx-reviewer/AGENTS.md >> ~/.codex/AGENTS.md
# AGENTS.md 内の SKILL_DIR を実際のパスに置換
sed -i "s|SKILL_DIR|$HOME/.codex/pptx-reviewer|g" ~/.codex/AGENTS.md
```

---

### STEP 4: python-pptx をインストールする

このスキルはスライドのテキスト抽出に Python ライブラリ `python-pptx` を使います。
以下のコマンドでインストールしてください（1 回だけ実行すれば OK）。

```cmd
pip install python-pptx
```

インストールできたか確認します。

```cmd
python -c "import pptx; print('インストール済みです')"
```

`インストール済みです` と表示されれば準備完了です。

---

### STEP 5: 実際に使ってみる

レビューしたい PPTX ファイルが入っているフォルダに移動してから Claude Code を起動します。

```cmd
cd C:\Users\（ユーザー名）\Documents
claude
```

チャット画面が開いたら、以下のように入力します。

```
proposal.pptx をレビューしてください。
```

Claude Code が自動でスキルを認識してレビューを開始します。
レビュー結果は **PPTX ファイルと同じフォルダ** に Markdown ファイルとして保存されます。

> **うまく動かない場合**：
> `%USERPROFILE%\.claude\skills\pptx-reviewer\SKILL.md` がエクスプローラーで確認できるか見てください。
> ファイルがなければ STEP 2 からやり直してください。

---

## Codex CLI（OpenAI）のセットアップ手順

> Claude Code CLI のセットアップ手順は[次のセクション](#はじめてのセットアップclaude-code-cli-導入手順)をご覧ください。

---

### STEP 1: Node.js のインストール確認

```cmd
node --version
```

バージョン番号（例: `v20.11.0`）が表示されれば OK です。
表示されない場合は [Node.js 公式サイト](https://nodejs.org/ja/) から **LTS 版**をインストールしてください。

---

### STEP 2: Python のインストール確認

```cmd
python --version
```

`Python 3.x.x` と表示されれば OK です。
表示されない場合は [Python 公式サイト](https://www.python.org/downloads/) からインストールしてください。

> **重要**：インストール時に **「Add Python to PATH」のチェックを必ずオンにしてください。**

---

### STEP 3: Codex CLI のインストール

```cmd
npm install -g @openai/codex
```

確認：

```cmd
codex --version
```

バージョン番号が表示されれば成功です。

---

### STEP 4: OpenAI API キーの取得と設定

#### 4-1. API キーを取得する

1. [OpenAI Platform](https://platform.openai.com/) にアクセスしてアカウントを作成・ログイン
2. 右上のアカウントメニュー → **「API keys」**
3. **「Create new secret key」** でキーを作成
4. 表示された `sk-...` で始まる文字列をコピーして保存

> **注意**：キーはこの画面を閉じると二度と表示されません。必ずすぐに保存してください。

#### 4-2. API キーを環境変数に設定する

```cmd
setx OPENAI_API_KEY "sk-ここにキーを貼り付ける"
```

設定後はコマンドプロンプトを**一度閉じて開き直します**。

確認：

```cmd
echo %OPENAI_API_KEY%
```

`sk-...` が表示されれば設定完了です。

---

### STEP 5: このスキルのインストールと配置

#### 5-1. リポジトリをダウンロードする

```cmd
git clone https://github.com/YuzuSuzu/pptx-review-text.git
```

#### 5-2. python-pptx をインストールする

```cmd
pip install python-pptx
```

確認：

```cmd
python -c "import pptx; print('OK')"
```

#### 5-3. スキルファイルを `~/.codex/` に配置する（推奨）

Codex CLI は **`C:\Users\（ユーザー名）\.codex\`** をグローバル設定フォルダとして使います。
ここに `AGENTS.md` を置くと、どのフォルダから `codex` を起動しても自動で読み込まれます。
Claude Code の `~/.claude/skills/` と同じ考え方です。

```cmd
:: .codex フォルダを作成（初回のみ）
mkdir "%USERPROFILE%\.codex"

:: リポジトリ全体を .codex\pptx-reviewer として配置
xcopy /E /I pptx-review-text "%USERPROFILE%\.codex\pptx-reviewer"

:: AGENTS.md を .codex 直下にコピー
copy "%USERPROFILE%\.codex\pptx-reviewer\AGENTS.md" "%USERPROFILE%\.codex\AGENTS.md"
```

配置後のフォルダ構成：

```
C:\Users\（ユーザー名）\
└── .codex\
    ├── AGENTS.md                        ← Codex CLI がここを自動で読み込む ★
    └── pptx-reviewer\                   ← スクリプト・用語リストの置き場
        ├── scripts\
        │   ├── extract_pptx.py
        │   ├── check_terminology.py
        │   └── create_dummy_pptx.py
        └── references\
            └── terminology.json
```

コピーした `%USERPROFILE%\.codex\AGENTS.md` をメモ帳で開き、`SKILL_DIR` を実際のパスに一括置換します。

```
変更前： SKILL_DIR
変更後： C:\Users\（ユーザー名）\.codex\pptx-reviewer
```

> **すでに `.codex\AGENTS.md` がある場合** は上書きせず、ファイルの末尾に内容を**追記**してください。

---

### STEP 6: 実際に使ってみる

```cmd
cd C:\Users\（ユーザー名）\pptx-review-text
codex
```

チャット画面が開いたら：

```
proposal.pptx をレビューしてください。
```

```
proposal.pptx の 2,5,6 ページをレビューしてください。
```

Codex CLI が `AGENTS.md` の手順に従い、スライドのテキスト抽出・用語チェック・レビュー分析を自動で実行し、Markdown レポートを保存します。

---

### よくあるトラブルと対処法（Codex CLI）

| 症状 | 原因と対処 |
|------|-----------|
| `codex: コマンドが見つかりません` | `npm install -g @openai/codex` を再実行。コマンドプロンプトを閉じて開き直す |
| `OPENAI_API_KEY is not set` | STEP 4-2 を確認。コマンドプロンプトを閉じて開き直す |
| スクリプトのパスエラー | `AGENTS.md` 内の `SKILL_DIR` が実際のパスに置き換えられているか確認 |
| `ModuleNotFoundError: No module named 'pptx'` | `pip install python-pptx` を実行 |
| AGENTS.md が読み込まれない | `codex` を起動しているフォルダに `AGENTS.md` があるか確認 |

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
