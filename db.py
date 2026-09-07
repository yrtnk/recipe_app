"""SQLite上のテーブル定義と保存処理（今回のスコープ: Book / Sheet / Formulation のみ）。

日時・ID・命名の方針:
  - ブックID・試作IDはExcel側（表紙・試作一覧の数式）で組み立てられたものを、
    そのまま正式なIDとしてDBに保存する（作成者コード＋ブック作成日時から自動生成）。
  - ブック作成日時はExcel表紙の自動計算セル（表紙!B7/D7）の表示用文字列を使う。
  - 別のブックと試作IDが衝突している場合はエラーとして検知する（DuplicateFormulationIdError）。
    ブックIDが「作成者コード＋分単位の作成日時」から決まるため、同じ人が同じ分に複数の
    ブックを新規作成した場合など、ごく稀に衝突しうる（そのための保険）。
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
    inspect,
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
    book_id = Column(String, unique=True, nullable=False)  # 表紙!B8（自動生成）
    project_id = Column(String)
    project_name = Column(String)
    category_code = Column(String)
    category_name = Column(String)
    brand_name = Column(String)
    created_by = Column(String)
    created_by_code = Column(String)
    book_created_date = Column(String)  # 表紙!B7（表示用文字列）
    book_created_time = Column(String)  # 表紙!D7（表示用文字列）
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
    formulation_id = Column(String)  # 試作ID（自動生成）
    title = Column(String)
    status = Column(String)
    created_by = Column(String)
    created_by_code = Column(String)
    trial_date = Column(String)  # 年/月/日から組み立てた表示用文字列
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
    """テーブルを作成する。既存のDBファイルが古いスキーマ（列が足りない等）のままだった場合は、
    テーブルを自動的に作り直す（開発中の頻繁なスキーマ変更に対応するため。今のところ試作データは
    使い捨て前提なので、既存データより「起動できること」を優先する）。
    全テーブル（books / sheets / formulations）を対象にチェックする。
    """
    inspector = inspect(_engine)
    needs_rebuild = False
    for table in Base.metadata.sorted_tables:
        if inspector.has_table(table.name):
            existing_cols = {c["name"] for c in inspector.get_columns(table.name)}
            expected_cols = {c.name for c in table.columns}
            if not expected_cols.issubset(existing_cols):
                needs_rebuild = True
                break
    if needs_rebuild:
        Base.metadata.drop_all(_engine)
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
    """別のブックに属するシートと試作IDが衝突している場合に送出する。"""

    def __init__(self, sheet_name: str, formulation_id: str, other_book_id: str, other_sheet_name: str):
        self.sheet_name = sheet_name
        self.formulation_id = formulation_id
        self.other_book_id = other_book_id
        self.other_sheet_name = other_sheet_name
        super().__init__(
            f"{sheet_name} の試作ID「{formulation_id}」は、別のブック「{other_book_id}」の"
            f"「{other_sheet_name}」で既に使われています。"
        )


def check_cross_book_duplicates(session, book_id_str: str, sheets: list) -> None:
    """アップロードしようとしているシートの試作IDが、他のブックで既に使われていないか確認する。
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
    別のブックと試作IDが衝突している場合は DuplicateFormulationIdError を送出する。
    """
    session = SessionLocal()
    try:
        check_cross_book_duplicates(session, parsed.book.book_id, parsed.sheets)

        book = session.query(Book).filter_by(book_id=parsed.book.book_id).one_or_none()
        if book is None:
            book = Book(book_id=parsed.book.book_id)
            session.add(book)

        book.project_id = parsed.book.project_id
        book.project_name = parsed.book.project_name
        book.category_code = parsed.book.category_code
        book.category_name = parsed.book.category_name
        book.brand_name = parsed.book.brand_name
        book.created_by = parsed.book.created_by
        book.created_by_code = parsed.book.created_by_code
        book.book_created_date = parsed.book.book_created_date
        book.book_created_time = parsed.book.book_created_time
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
            sheet.trial_date = s.trial_date
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
