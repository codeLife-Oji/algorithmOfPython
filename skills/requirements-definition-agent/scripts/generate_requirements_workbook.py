"""要件定義JSONから要件一覧（Excel・6シート）を生成する。

使い方:
    python3 generate_requirements_workbook.py requirements.json /mnt/user-data/outputs/要件一覧.xlsx

シート構成:
    ① 業務要件   ② 機能要件   ③ 非機能要件
    ④ 制約・前提  ⑤ スコープ外  ⑥ 確認事項

JSONスキーマは requirements_schema.py のdocstringを参照。
非機能要件シートは、言及の無かった分類も「未確定」の行として必ず出力する
（行が無い分類は、誰も見落としに気づけないため）。
"""

import os
import sys

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import requirements_schema as schema  # noqa: E402

HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT = Font(color="FFFFFF", bold=True, size=10)
BODY_FONT = Font(size=10)
ALERT_FILL = PatternFill("solid", fgColor="FFF2CC")  # 未確定行の強調
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

# (シート名, 列定義[(見出し, JSONキー, 幅)], JSONのセクションキー)
SHEETS = [
    ("①業務要件", [
        ("ID", "id", 10), ("業務分類", "category", 16), ("業務要件", "title", 34),
        ("詳細", "detail", 50), ("主担当（アクター）", "actor", 16),
        ("優先度", "priority", 10), ("状態", "status", 10), ("根拠・情報源", "source", 26),
    ], "business_requirements"),

    ("②機能要件", [
        ("ID", "id", 10), ("大分類", "major", 16), ("中分類", "minor", 16),
        ("機能名", "name", 26), ("機能概要", "detail", 50),
        ("アクター", "actor", 16), ("関連業務要件", "related_br", 14),
        ("優先度", "priority", 10), ("状態", "status", 10), ("根拠・情報源", "source", 24),
    ], "functional_requirements"),

    ("③非機能要件", [
        ("ID", "id", 10), ("分類", "category", 20), ("確認項目", "item", 24),
        ("要件値（数値と条件）", "value", 46), ("根拠・背景", "basis", 34),
        ("状態", "status", 10), ("関連確認事項", "question_id", 14),
    ], "non_functional_requirements"),

    ("④制約・前提", [
        ("ID", "id", 10), ("区分", "type", 10), ("内容", "detail", 50),
        ("崩れた場合の影響", "impact", 40), ("根拠・情報源", "source", 24),
    ], "constraints"),

    ("⑤スコープ外", [
        ("ID", "id", 10), ("対象", "detail", 40), ("理由", "reason", 34),
        ("備考", "note", 20),
    ], "out_of_scope"),

    ("⑥確認事項", [
        ("ID", "id", 10), ("確認先", "target", 12), ("確認内容", "detail", 54),
        ("背景・確認理由", "why", 40), ("関連要件", "related", 14),
        ("回答期限", "due", 14), ("状態", "status", 10),
    ], "questions"),
]


def fill_nfr_gaps(rows):
    """非機能要件6分類のうち、1行も無い分類を「未確定」の行として補う。"""
    present = {r.get("category") for r in rows}
    filled = list(rows)
    for i, category in enumerate(schema.NFR_CATEGORIES, start=1):
        if category not in present:
            filled.append({
                "id": f"NFR-X{i:02d}",
                "category": category,
                "item": "（言及なし）",
                "value": "未確定",
                "basis": "ヒアリングで言及がなかったため要確認",
                "status": "未確定",
                "question_id": "",
            })
    order = {c: i for i, c in enumerate(schema.NFR_CATEGORIES)}
    filled.sort(key=lambda r: (order.get(r.get("category"), 99), str(r.get("id"))))
    return filled


def is_unresolved(row):
    return str(row.get("status", "")) in ("未確定", "未回答", "仮置き")


def write_sheet(ws, columns, rows):
    ws.freeze_panes = "A2"
    for col, (header, _, width) in enumerate(columns, start=1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.fill, cell.font, cell.border = HEADER_FILL, HEADER_FONT, BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(col)].width = width
    ws.row_dimensions[1].height = 24

    for r, row in enumerate(rows, start=2):
        highlight = is_unresolved(row)
        for col, (_, key, _) in enumerate(columns, start=1):
            cell = ws.cell(row=r, column=col, value=row.get(key, ""))
            cell.font, cell.border = BODY_FONT, BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            if highlight:
                cell.fill = ALERT_FILL

    if rows:
        ws.auto_filter.ref = (
            f"A1:{get_column_letter(len(columns))}{len(rows) + 1}"
        )


def write_cover(ws, data, counts):
    project = data["project"]
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 60
    rows = [
        ("プロジェクト名", project["name"]),
        ("顧客", project["client"]),
        ("版数", project["version"]),
        ("作成日", project["date"]),
        ("作成者", project["author"]),
        ("", ""),
        ("業務要件", len(data["business_requirements"])),
        ("機能要件", len(data["functional_requirements"])),
        ("非機能要件", len(data["non_functional_requirements"])),
        ("制約・前提", len(data["constraints"])),
        ("スコープ外", len(data["out_of_scope"])),
        ("確認事項", len(data["questions"])),
        ("", ""),
        ("未確定の非機能要件", counts["nfr_unresolved"]),
        ("未回答の確認事項", counts["questions_open"]),
        ("仮置きの要件", counts["provisional"]),
        ("", ""),
        ("情報源", " / ".join(str(s) for s in data["sources"])),
    ]
    for r, (label, value) in enumerate(rows, start=1):
        label_cell = ws.cell(row=r, column=1, value=label)
        label_cell.font = Font(bold=True, size=10)
        value_cell = ws.cell(row=r, column=2, value=value)
        value_cell.font = BODY_FONT
        value_cell.alignment = Alignment(vertical="top", wrap_text=True)


def main():
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(1)

    data = schema.load(sys.argv[1])
    out_path = schema.prepare_output(sys.argv[2])
    counts = schema.unresolved_counts(data)

    wb = Workbook()
    write_cover(wb.active, data, counts)
    wb.active.title = "表紙・サマリ"

    for title, columns, key in SHEETS:
        rows = data[key]
        if key == "non_functional_requirements":
            rows = fill_nfr_gaps(rows)
        write_sheet(wb.create_sheet(title), columns, rows)

    wb.save(out_path)
    print(f"生成しました: {out_path}")
    print(
        f"未確定の非機能要件 {counts['nfr_unresolved']}件 / "
        f"未回答の確認事項 {counts['questions_open']}件 / "
        f"仮置きの要件 {counts['provisional']}件"
    )


if __name__ == "__main__":
    main()
