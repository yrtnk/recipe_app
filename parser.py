"""
recipe_template.xlsx（表紙・コード表・試作シートN・試作一覧）を読み取り、
構造化データ（book / sheets / formulations）に変換するパーサー。

今回のスコープ:
  - 配合表（原料・配合量）はまだ扱わない
  - 試作一覧シートの内容（シート単位のメタデータ）
  - 各試作シートの13〜17行目（配合No・配合タイトル・配合目的意図・配合結果・試作評価）

前提:
  - このファイルは Excel（または recalc 済み）で一度保存されていること。
    openpyxl は data_only=True で「計算済みの値」を読むため、保存時に
    数式が再計算されていないファイルは値が None で返ってくる。
"""
from __future__ import annotations

import datetime
from dataclasses import dataclass, field
from typing import Optional

import openpyxl
from openpyxl.utils import get_column_letter

PAIR_START_COLS = [3 + 2 * i for i in range(12)]  # C,E,G,...,Y (12配合)
FORMULATION_META_ROWS = {
    "formulation_no": 13,
    "formulation_title": 14,
    "purpose_intent": 15,
    "result": 16,
    "evaluation": 17,
}


@dataclass
class BookData:
    book_id: Optional[str]
    book_name: Optional[str]
    project_id: Optional[str]
    category_code: Optional[str]
    created_by: Optional[str]
    created_by_code: Optional[str]
    created_date: Optional[datetime.date]
    created_time: Optional[datetime.time]


@dataclass
class SheetData:
    trial_no: Optional[int]
    sheet_name: Optional[str]
    formulation_id: Optional[str]
    title: Optional[str]
    status: Optional[str]
    created_by: Optional[str]
    created_by_code: Optional[str]
    created_date: Optional[datetime.date]
    created_time: Optional[datetime.time]
    updated_date: Optional[datetime.date]
    updated_time: Optional[datetime.time]
    completed_date: Optional[datetime.date]
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
        book_id=ws["B3"].value,
        book_name=ws["D3"].value,
        project_id=ws["B4"].value,
        category_code=ws["D4"].value,
        created_by=ws["B5"].value,
        created_by_code=ws["D5"].value,
        created_date=ws["B6"].value,
        created_time=ws["D6"].value,
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
                warnings.append(f"{sheet_name}: 配合ID未確定（作成者コード/作成日時が未入力）")
            sheets.append(
                SheetData(
                    trial_no=int(trial_no) if trial_no is not None else None,
                    sheet_name=sheet_name,
                    formulation_id=formulation_id,
                    title=ws.cell(row=row, column=4).value,
                    status=ws.cell(row=row, column=5).value,
                    created_by=ws.cell(row=row, column=6).value,
                    created_by_code=ws.cell(row=row, column=7).value,
                    created_date=ws.cell(row=row, column=8).value,
                    created_time=ws.cell(row=row, column=9).value,
                    updated_date=ws.cell(row=row, column=10).value,
                    updated_time=ws.cell(row=row, column=11).value,
                    completed_date=ws.cell(row=row, column=12).value,
                    background_purpose=ws.cell(row=row, column=13).value,
                    conclusion=ws.cell(row=row, column=14).value,
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


def parse_workbook(path: str) -> ParsedWorkbook:
    wb = openpyxl.load_workbook(path, data_only=True)
    warnings: list[str] = []

    book = parse_book(wb)
    sheets = parse_sheets(wb, warnings)

    formulations: list[FormulationData] = []
    for sheet in sheets:
        if sheet.sheet_name not in wb.sheetnames:
            warnings.append(f"{sheet.sheet_name}: 一覧には存在するが、対応するシートが見つかりません")
            continue
        ws = wb[sheet.sheet_name]
        sheet_formulations = parse_formulations_for_sheet(ws, sheet.sheet_name)
        if not sheet_formulations:
            warnings.append(f"{sheet.sheet_name}: 配合タイトルが入力された配合が見つかりません")
        formulations.extend(sheet_formulations)

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
