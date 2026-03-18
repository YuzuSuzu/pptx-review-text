"""
用語統一リスト（terminology.json）を使って、
extract_pptx.py が出力したJSONに表記ゆれがないかチェックするスクリプト。

Usage:
  python check_terminology.py <extracted-json> <terminology-json>

例:
  python check_terminology.py extract_out.json references/terminology.json

出力: JSON形式で発見した表記ゆれをstdoutに出力する
"""
import sys
import json
import re
import argparse


def load_json(path):
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
    parser = argparse.ArgumentParser(
        description="抽出済みJSONと用語統一リストを照合して表記ゆれを検出する"
    )
    parser.add_argument("extracted_json", help="extract_pptx.py の出力JSONファイル")
    parser.add_argument("terminology_json", help="references/terminology.json のパス")
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
        sys.exit(1)

    terms = terminology.get("terms", [])
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

    output = {
        "total_slides_checked": len(extracted.get("slides", [])),
        "slides_with_issues": len(results),
        "terminology_version": terminology.get("version", "unknown"),
        "results": results,
    }
    import io
    stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
    stdout.write(json.dumps(output, ensure_ascii=False, indent=2))
    stdout.flush()


if __name__ == "__main__":
    main()
