"""
用語統一リスト（terminology.json）を使って、
extract_pptx.py が出力したJSONに表記ゆれがないかチェックするスクリプト。

terminology_json 引数は省略可能。省略時はこのスクリプトと同じスキルフォルダ内の
references/terminology.json を自動的に参照する。

Usage:
  # パイプで直接渡す場合（中間ファイル不要・推奨）
  python extract_pptx.py proposal.pptx | python check_terminology.py -

  # ファイルを経由する場合
  python check_terminology.py extract_out.json

  # 用語リストを明示指定する場合
  python check_terminology.py - /path/to/terminology.json

出力: JSON形式で発見した表記ゆれをstdoutに出力する
"""
import sys
import io
import os
import json
import re
import argparse


def _setup_utf8():
    """stdout/stderr を UTF-8 に統一する（Windows の cp932 対策）。"""
    os.environ.setdefault("PYTHONUTF8", "1")
    if hasattr(sys.stdout, "buffer"):
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    if hasattr(sys.stderr, "buffer"):
        sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")


def load_json(path):
    """パスが "-" の場合は stdin から読み込む。"""
    if path == "-":
        stdin = io.TextIOWrapper(sys.stdin.buffer, encoding="utf-8")
        return json.load(stdin)
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def find_variants_in_text(text, correct, variants):
    """
    テキスト中にvariantsが含まれているか検索し、
    見つかった箇所（文脈つき）を返す。
    """
    hits = []
    for variant in variants:
        # 大文字小文字を区別してマッチ（日本語は区別なし）
        pattern = re.compile(re.escape(variant))
        for m in pattern.finditer(text):
            # 前後20文字のコンテキストを取得
            start = max(0, m.start() - 20)
            end = min(len(text), m.end() + 20)
            context = text[start:end].replace("\n", " ")
            hits.append({
                "found": variant,
                "correct": correct,
                "context": f"…{context}…",
            })
    return hits


def main():
    _setup_utf8()

    # このスクリプトの場所から references/terminology.json を自動検出
    _script_dir = os.path.dirname(os.path.abspath(__file__))
    _default_terminology = os.path.join(_script_dir, "..", "references", "terminology.json")

    parser = argparse.ArgumentParser(
        description="抽出済みJSONと用語統一リストを照合して表記ゆれを検出する"
    )
    parser.add_argument("extracted_json", help="extract_pptx.py の出力JSONファイル（'-' で stdin）")
    parser.add_argument(
        "terminology_json",
        nargs="?",
        default=_default_terminology,
        help=f"references/terminology.json のパス（省略時: {_default_terminology}）",
    )
    args = parser.parse_args()

    try:
        extracted = load_json(args.extracted_json)
    except Exception as e:
        print(f"ERROR: 抽出JSONを読み込めませんでした: {e}", file=sys.stderr)
        sys.exit(1)

    try:
        terminology = load_json(args.terminology_json)
    except Exception as e:
        print(f"ERROR: 用語リストを読み込めませんでした: {e}", file=sys.stderr)
        print(f"       パス: {args.terminology_json}", file=sys.stderr)
        sys.exit(1)

    terms = terminology.get("terms", [])
    print(f"[2/2] 用語チェック中... ({len(terms)}語)", file=sys.stderr)
    results = []

    for slide in extracted.get("slides", []):
        slide_number = slide["slide_number"]
        slide_title = slide["title"]
        slide_hits = []

        # スライド内の全テキストを結合して検索
        for shape in slide.get("shapes", []):
            for para in shape.get("paragraphs", []):
                text = para.get("text", "")
                for term in terms:
                    correct = term["correct"]
                    variants = term.get("variants", [])
                    hits = find_variants_in_text(text, correct, variants)
                    for h in hits:
                        h["shape"] = shape["shape_name"]
                        slide_hits.append(h)

        if slide_hits:
            results.append({
                "slide_number": slide_number,
                "slide_title": slide_title,
                "hits": slide_hits,
            })

    total_hits = sum(len(r["hits"]) for r in results)
    print(f"[2/2] 用語チェック完了 — {len(results)}スライドで{total_hits}件の表記ゆれを検出", file=sys.stderr)

    output = {
        "total_slides_checked": len(extracted.get("slides", [])),
        "slides_with_issues": len(results),
        "terminology_version": terminology.get("version", "unknown"),
        "results": results,
    }
    sys.stdout.write(json.dumps(output, ensure_ascii=False, indent=2))
    sys.stdout.flush()


if __name__ == "__main__":
    main()
