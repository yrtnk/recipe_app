"""SQLite上のテーブル定義と保存処理（今回のスコープ: Book / Sheet / Formulation のみ）。

日時・配合IDの方針:
  - 配合ID・作成日時はExcel側（表紙・試作シート・試作一覧の数式）で組み立てられたものを、
    そのまま正式なIDとしてDBに保存する。
  - ブック単位の作成日時はExcelファイル自体のプロパティ（parser.BookData.file_created_at）を使う。
  - 別のブックと配合IDが衝突している場合はエラーとして検知する（DuplicateFormulationIdError）。
"""
from __future__ import annotations

import datetime

from sqlalchemy import (
    Column,
    DateTime,
    ForeignKey,
    Integer,
    String,
    Text,
    UniqueConstraint,
    create_engine,
)
from sqlalchemy.orm import declarative_base, relationship, sessionmaker

from parser import ParsedWorkbook

Base = declarative_base()

DB_PATH = "sqlite:///recipe.db"
_engine = create_engine(DB_PATH, future=True)
SessionLocal = sessionmaker(bind=_engine, future=True)


class Book(Base):
    __tablename__ = "books"

    id = Column(Integer, primary_key=True, autoincrement=True)
    book_id = Column(String, unique=True, nullable=False)
    book_name = Column(String)
    project_id = Column(String)
    category_code = Column(String)
    created_by = Column(String)
    created_by_code = Column(String)
    file_created_at = Column(DateTime)  # Excelファイル自体の作成日時
    source_filename = Column(String)
    imported_at = Column(DateTime, default=datetime.datetime.utcnow)  # 最後にアップロードされた日時

    sheets = relationship("Sheet", back_populates="book", cascade="all, delete-orphan")


class Sheet(Base):
    __tablename__ = "sheets"
    __table_args__ = (UniqueConstraint("book_id", "sheet_name", name="uq_book_sheet"),)

    id = Column(Integer, primary_key=True, autoincrement=True)
    book_id = Column(Integer, ForeignKey("books.id"), nullable=False)
    trial_no = Column(Integer)
    sheet_name = Column(String, nullable=False)
    formulation_id = Column(String)
    title = Column(String)
    status = Column(String)
    created_by = Column(String)
    created_by_code = Column(String)
    created_date = Column(String)
    created_time = Column(String)
    updated_date = Column(String)
    updated_time = Column(String)
    completed_date = Column(String)
    background_purpose = Column(Text)
    conclusion = Column(Text)

    book = relationship("Book", back_populates="sheets")
    formulations = relationship("Formulation", back_populates="sheet", cascade="all, delete-orphan")


class Formulation(Base):
    __tablename__ = "formulations"

    id = Column(Integer, primary_key=True, autoincrement=True)
    sheet_id = Column(Integer, ForeignKey("sheets.id"), nullable=False)
    formulation_no = Column(String)
    formulation_title = Column(String)
    purpose_intent = Column(Text)
    result = Column(Text)
    evaluation = Column(Text)

    sheet = relationship("Sheet", back_populates="formulations")


def init_db():
    Base.metadata.create_all(_engine)


def get_known_user_codes() -> list[str]:
    """DBに登録済みの作成者コード一覧（ログインユーザー選択用）。"""
    session = SessionLocal()
    try:
        codes = {b.created_by_code for b in session.query(Book.created_by_code).distinct() if b.created_by_code}
        codes |= {s.created_by_code for s in session.query(Sheet.created_by_code).distinct() if s.created_by_code}
        return sorted(codes)
    finally:
        session.close()


class DuplicateFormulationIdError(Exception):
    """別のブックに属するシートと配合IDが衝突している場合に送出する。"""

    def __init__(self, sheet_name: str, formulation_id: str, other_book_id: str, other_sheet_name: str):
        self.sheet_name = sheet_name
        self.formulation_id = formulation_id
        self.other_book_id = other_book_id
        self.other_sheet_name = other_sheet_name
        super().__init__(
            f"{sheet_name} の配合ID「{formulation_id}」は、別のブック「{other_book_id}」の"
            f"「{other_sheet_name}」で既に使われています。作成日・作成時刻をご確認ください。"
        )


def _to_str(v):
    """openpyxlが返す日付/時刻オブジェクトを表示用の文字列に統一する。"""
    if v is None:
        return None
    if isinstance(v, (datetime.date, datetime.datetime, datetime.time)):
        return str(v)
    return str(v)


def check_cross_book_duplicates(session, book_id_str: str, sheets: list) -> None:
    """アップロードしようとしているシートの配合IDが、他のブックで既に使われていないか確認する。
    プレースホルダー「(入力待ち)」や空欄は対象外（未確定なので重複扱いしない）。
    """
    for s in sheets:
        fid = s.formulation_id
        if not fid or fid == "(入力待ち)":
            continue
        existing = (
            session.query(Sheet, Book)
            .join(Book, Sheet.book_id == Book.id)
            .filter(Sheet.formulation_id == fid, Book.book_id != book_id_str)
            .first()
        )
        if existing:
            existing_sheet, existing_book = existing
            raise DuplicateFormulationIdError(
                sheet_name=s.sheet_name,
                formulation_id=fid,
                other_book_id=existing_book.book_id,
                other_sheet_name=existing_sheet.sheet_name,
            )


def save_parsed_workbook(parsed: ParsedWorkbook, source_filename: str) -> dict:
    """パース結果をDBに保存する（book_idが既存なら更新、シート/配合は洗い替え）。
    別のブックと配合IDが衝突している場合は DuplicateFormulationIdError を送出する。
    """
    session = SessionLocal()
    try:
        check_cross_book_duplicates(session, parsed.book.book_id, parsed.sheets)

        book = session.query(Book).filter_by(book_id=parsed.book.book_id).one_or_none()
        if book is None:
            book = Book(book_id=parsed.book.book_id)
            session.add(book)

        book.book_name = parsed.book.book_name
        book.project_id = parsed.book.project_id
        book.category_code = parsed.book.category_code
        book.created_by = parsed.book.created_by
        book.created_by_code = parsed.book.created_by_code
        book.file_created_at = parsed.book.file_created_at
        book.source_filename = source_filename
        book.imported_at = datetime.datetime.utcnow()
        session.flush()  # book.id を確定させる

        formulations_by_sheet: dict[str, list] = {}
        for f in parsed.formulations:
            formulations_by_sheet.setdefault(f.sheet_name, []).append(f)

        sheet_count = 0
        formulation_count = 0

        for s in parsed.sheets:
            sheet = (
                session.query(Sheet)
                .filter_by(book_id=book.id, sheet_name=s.sheet_name)
                .one_or_none()
            )
            if sheet is None:
                sheet = Sheet(book_id=book.id, sheet_name=s.sheet_name)
                session.add(sheet)

            sheet.trial_no = s.trial_no
            sheet.formulation_id = s.formulation_id
            sheet.title = s.title
            sheet.status = s.status
            sheet.created_by = s.created_by
            sheet.created_by_code = s.created_by_code
            sheet.created_date = _to_str(s.created_date)
            sheet.created_time = _to_str(s.created_time)
            sheet.updated_date = _to_str(s.updated_date)
            sheet.updated_time = _to_str(s.updated_time)
            sheet.completed_date = _to_str(s.completed_date)
            sheet.background_purpose = s.background_purpose
            sheet.conclusion = s.conclusion
            session.flush()  # sheet.id を確定させる
            sheet_count += 1

            # このシートの配合は洗い替え（毎回全部作り直す）
            session.query(Formulation).filter_by(sheet_id=sheet.id).delete()
            for f in formulations_by_sheet.get(s.sheet_name, []):
                session.add(
                    Formulation(
                        sheet_id=sheet.id,
                        formulation_no=f.formulation_no,
                        formulation_title=f.formulation_title,
                        purpose_intent=f.purpose_intent,
                        result=f.result,
                        evaluation=f.evaluation,
                    )
                )
                formulation_count += 1

        session.commit()
        return {"book_id": book.book_id, "sheets": sheet_count, "formulations": formulation_count}
    except Exception:
        session.rollback()
        raise
    finally:
        session.close()
