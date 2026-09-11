"""要件定義JSONの共通スキーマ定義とローダ。

Word生成・Excel生成の両スクリプトから読み込まれ、同じJSONから同じ内容が出ることを保証する。

JSONスキーマ（すべてのキーは任意。無い場合は空として扱う）:

{
  "project": {
    "name": "受注管理システム再構築",
    "client": "株式会社サンプル",
    "date": "2026-09-11",
    "version": "0.1",
    "author": "尾嶋",
    "purpose": "背景と目的の本文（複数段落は \\n\\n 区切り）"
  },
  "sources": ["2026-09-03 キックオフ議事録", "現行受注管理表.xlsx"],
  "glossary": [{"term": "受注", "definition": "..."}],

  "business_requirements": [{
    "id": "BR-001", "category": "受注管理", "title": "受注入力の二重作業をなくす",
    "detail": "...", "actor": "営業担当", "priority": "Must",
    "status": "確定", "source": "9/3 田中様 発言"
  }],

  "functional_requirements": [{
    "id": "FR-001", "major": "受注管理", "minor": "受注登録", "name": "受注登録画面",
    "detail": "...", "actor": "営業担当", "related_br": "BR-001",
    "priority": "Must", "status": "確定", "source": "9/3 議事録"
  }],

  "non_functional_requirements": [{
    "id": "NFR-001", "category": "性能・拡張性", "item": "オンライン応答時間",
    "value": "3秒以内（95%tile / 同時100ユーザ時）", "basis": "...",
    "status": "確定", "question_id": ""
  }],

  "constraints": [{
    "id": "CON-001", "type": "制約", "detail": "...", "impact": "...", "source": "..."
  }],

  "out_of_scope": [{
    "id": "OUT-001", "detail": "多言語対応", "reason": "フェーズ2で実施", "note": "未確認"
  }],

  "questions": [{
    "id": "Q-001", "target": "顧客", "detail": "...", "why": "...",
    "related": "NFR-001", "due": "2026-09-20", "status": "未回答"
  }],

  "revisions": [{"version": "0.1", "date": "2026-09-11", "note": "初版", "author": "尾嶋"}]
}

非機能要件の6分類（この順で出力する）:
  可用性 / 性能・拡張性 / 運用・保守性 / 移行性 / セキュリティ / システム環境・エコロジー
"""

import json
import os
import sys

NFR_CATEGORIES = [
    "可用性",
    "性能・拡張性",
    "運用・保守性",
    "移行性",
    "セキュリティ",
    "システム環境・エコロジー",
]

SECTIONS = {
    "business_requirements": "業務要件",
    "functional_requirements": "機能要件",
    "non_functional_requirements": "非機能要件",
    "constraints": "制約・前提条件",
    "out_of_scope": "スコープ外",
    "questions": "確認事項",
}


def load(path):
    """要件定義JSONを読み込み、欠けているキーを空で埋めて返す。"""
    with open(path, encoding="utf-8") as f:
        data = json.load(f)

    data.setdefault("project", {})
    data["project"].setdefault("name", "（プロジェクト名未設定）")
    for key in ("client", "date", "version", "author", "purpose"):
        data["project"].setdefault(key, "")

    for key in list(SECTIONS) + ["sources", "glossary", "revisions"]:
        data.setdefault(key, [])

    return data


def prepare_output(path):
    """出力先ディレクトリを用意する。書けない場合はカレント配下にフォールバックする。"""
    directory = os.path.dirname(os.path.abspath(path))
    try:
        os.makedirs(directory, exist_ok=True)
        return path
    except OSError:
        fallback = os.path.join(os.getcwd(), os.path.basename(path))
        print(f"警告: {directory} に書けないため {fallback} に出力します", file=sys.stderr)
        return fallback


def nfr_by_category(data):
    """非機能要件を6分類の順に並べ、言及の無い分類も空リストで返す。

    行が存在しない分類は誰も気づけないため、呼び出し側が「未確定」として
    明示できるよう、必ず6分類すべてをキーとして返す。
    """
    grouped = {category: [] for category in NFR_CATEGORIES}
    for row in data["non_functional_requirements"]:
        category = row.get("category", "")
        grouped.setdefault(category, []).append(row)
    return grouped


def unresolved_counts(data):
    """未確定・未回答の件数を集計する。チャットへの要約に使う。"""
    nfr_open = [
        r for r in data["non_functional_requirements"]
        if r.get("status") not in ("確定", "")
    ]
    q_open = [q for q in data["questions"] if q.get("status") != "回答済"]
    provisional = [
        r for r in data["business_requirements"] + data["functional_requirements"]
        if r.get("status") == "仮置き"
    ]
    return {
        "nfr_unresolved": len(nfr_open),
        "questions_open": len(q_open),
        "provisional": len(provisional),
    }
