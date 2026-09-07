"""
recipe_template.xlsx（使い方・コード表・表紙・試作一覧・試作1〜30）を読み取り、
構造化データ（book / sheets / formulations）に変換するパーサー。

今回のスコープ:
  - 配合表（原料・配合量・パーツ・画像貼付け欄・試作条件・測定結果）はまだ扱わない
  - 試作一覧シートの内容（シート単位のメタデータ）
  - 各試作シートの9〜13行目（配合No・配合タイトル・配合目的意図・配合結果・試作評価）

表紙のレイアウト:
  B3=企画ID  D3=企画名
  B4=カテゴリーコード  D4=カテゴリー名
  B5=作成者  D5=作成者コード
  B6=ブランド名
  B7=ブック作成日（自動計算）  D7=ブック作成時刻（自動計算）
  B8=ブックID（自動計算。カテゴリーコードではなく作成者コード＋ブック作成日時から組み立てる）

試作Nシートのレイアウト:
  B1=試作ID（自動計算）  F1=ステータス
  B2=タイトル
  B3=作成者  D3=作成者コード（自動計算）  F3/G3/H3=試作日（年/月/日）  I3=試作日（組み立て済み表示用）
  B4=背景・目的
  B5=結論
  B6=次のアクション
  9〜13行目=配合No/配合タイトル/配合目的意図/配合結果/試作評価（配合列はC,E,G,...Yの12列）

試作一覧のレイアウト（列）:
  A=試作No  B=シート名  C=試作ID  D=タイトル  E=ステータス  F=作成者  G=作成者コード
  H=試作日（年/月/日から組み立てた表示用文字列）  I=背景・目的  J=結論

前提:
  - このファイルは Excel（または recalc 済み）で一度保存されていること。
    openpyxl は data_only=True で「計算済みの値」を読むため、保存時に
    数式が再計算されていないファイルは値が None で返ってくる。
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Optional

import openpyxl
from openpyxl.utils import get_column_letter

PAIR_START_COLS = [3 + 2 * i for i in range(12)]  # C,E,G,...,Y (12配合)
FORMULATION_META_ROWS = {
    "formulation_no": 9,
    "formulation_title": 10,
    "purpose_intent": 11,
    "result": 12,
    "evaluation": 13,
}


@dataclass
class BookData:
    book_id: Optional[str]  # 表紙!B8（自動生成）
    project_id: Optional[str]
    project_name: Optional[str]
    category_code: Optional[str]
    category_name: Optional[str]
    brand_name: Optional[str]
    created_by: Optional[str]
    created_by_code: Optional[str]
    book_created_date: Optional[str]  # 表紙!B7（表示用文字列）
    book_created_time: Optional[str]  # 表紙!D7（表示用文字列）


@dataclass
class SheetData:
    trial_no: Optional[int]
    sheet_name: Optional[str]
    formulation_id: Optional[str]  # 試作ID（自動生成）
    title: Optional[str]
    status: Optional[str]
    created_by: Optional[str]
    created_by_code: Optional[str]
    trial_date: Optional[str]  # 年/月/日から組み立てた表示用文字列（例: 2026/07/02）
    background_purpose: Optional[str]
    conclusion: Optional[str]


@dataclass
class FormulationData:
    sheet_name: str
    formulation_no: Optional[str]
    formulation_title: Optional[str]
    purpose_intent: Optional[str]
    result: Optional[str]
    evaluation: Optional[str]


@dataclass
class ParsedWorkbook:
    book: BookData
    sheets: list[SheetData] = field(default_factory=list)
    formulations: list[FormulationData] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)


def _blank(v):
    return v is None or (isinstance(v, str) and v.strip() == "")


def parse_book(wb) -> BookData:
    ws = wb["表紙"]
    return BookData(
        book_id=ws["B8"].value,
        project_id=ws["B3"].value,
        project_name=ws["D3"].value,
        category_code=ws["B4"].value,
        category_name=ws["D4"].value,
        brand_name=ws["B6"].value,
        created_by=ws["B5"].value,
        created_by_code=ws["D5"].value,
        book_created_date=ws["B7"].value,
        book_created_time=ws["D7"].value,
    )


def parse_sheets(wb, warnings: list[str]) -> list[SheetData]:
    ws = wb["試作一覧"]
    sheets: list[SheetData] = []
    row = 2
    while True:
        sheet_name = ws.cell(row=row, column=2).value  # B列: シート名
        trial_no = ws.cell(row=row, column=1).value  # A列: 試作No
        if trial_no is None:
            break  # 一覧の行が尽きた
        if not _blank(sheet_name):
            formulation_id = ws.cell(row=row, column=3).value
            if formulation_id == "(入力待ち)":
                warnings.append(f"{sheet_name}: 試作ID未確定（表紙の作成者コードが未入力）")
            sheets.append(
                SheetData(
                    trial_no=int(trial_no) if trial_no is not None else None,
                    sheet_name=sheet_name,
                    formulation_id=formulation_id,
                    title=ws.cell(row=row, column=4).value,
                    status=ws.cell(row=row, column=5).value,
                    created_by=ws.cell(row=row, column=6).value,
                    created_by_code=ws.cell(row=row, column=7).value,
                    trial_date=ws.cell(row=row, column=8).value,
                    background_purpose=ws.cell(row=row, column=9).value,
                    conclusion=ws.cell(row=row, column=10).value,
                )
            )
        row += 1
    return sheets


def parse_formulations_for_sheet(ws, sheet_name: str) -> list[FormulationData]:
    results = []
    for start_col in PAIR_START_COLS:
        col_letter = get_column_letter(start_col)
        title = ws[f"{col_letter}{FORMULATION_META_ROWS['formulation_title']}"].value
        if _blank(title):
            continue  # このスロットは未使用の配合とみなしてスキップ
        results.append(
            FormulationData(
                sheet_name=sheet_name,
                formulation_no=ws[f"{col_letter}{FORMULATION_META_ROWS['formulation_no']}"].value,
                formulation_title=title,
                purpose_intent=ws[f"{col_letter}{FORMULATION_META_ROWS['purpose_intent']}"].value,
                result=ws[f"{col_letter}{FORMULATION_META_ROWS['result']}"].value,
                evaluation=ws[f"{col_letter}{FORMULATION_META_ROWS['evaluation']}"].value,
            )
        )
    return results


def _sheet_is_touched(sheet: "SheetData", has_formulations: bool) -> bool:
    """デフォルト値のまま（何も入力されていない）試作シートかどうかを判定する。
    ステータスは常に「作成中」が初期値として入っているため、判定対象には含めない。
    """
    fields = [
        sheet.title,
        sheet.created_by,
        sheet.created_by_code,
        sheet.trial_date,
        sheet.background_purpose,
        sheet.conclusion,
    ]
    if any(not _blank(v) for v in fields):
        return True
    return has_formulations


def parse_workbook(path: str) -> ParsedWorkbook:
    wb = openpyxl.load_workbook(path, data_only=True)
    warnings: list[str] = []

    book = parse_book(wb)
    all_sheets = parse_sheets(wb, warnings)

    formulations_by_sheet: dict[str, list[FormulationData]] = {}
    for sheet in all_sheets:
        if sheet.sheet_name not in wb.sheetnames:
            warnings.append(f"{sheet.sheet_name}: 一覧には存在するが、対応するシートが見つかりません")
            continue
        ws = wb[sheet.sheet_name]
        sheet_formulations = parse_formulations_for_sheet(ws, sheet.sheet_name)
        # まったくの未使用シート（タイトルも未入力）は警告の対象外にする（30枚あらかじめ用意しているため）
        if not sheet_formulations and not _blank(sheet.title):
            warnings.append(f"{sheet.sheet_name}: 配合タイトルが入力された配合が見つかりません")
        formulations_by_sheet[sheet.sheet_name] = sheet_formulations

    # デフォルト値のまま（何も入力されていない）試作シートは、アップロード対象から除外する
    sheets = [
        s for s in all_sheets
        if _sheet_is_touched(s, bool(formulations_by_sheet.get(s.sheet_name)))
    ]
    formulations: list[FormulationData] = []
    for s in sheets:
        formulations.extend(formulations_by_sheet.get(s.sheet_name, []))

    return ParsedWorkbook(book=book, sheets=sheets, formulations=formulations, warnings=warnings)


if __name__ == "__main__":
    import sys

    path = sys.argv[1] if len(sys.argv) > 1 else "sample_filled.xlsx"
    parsed = parse_workbook(path)
    print("=== BOOK ===")
    print(parsed.book)
    print(f"=== SHEETS ({len(parsed.sheets)}) ===")
    for s in parsed.sheets:
        print(s)
    print(f"=== FORMULATIONS ({len(parsed.formulations)}) ===")
    for f in parsed.formulations:
        print(f)
    if parsed.warnings:
        print("=== WARNINGS ===")
        for w in parsed.warnings:
            print("-", w)
