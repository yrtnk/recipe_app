import tempfile

import pandas as pd
import streamlit as st

from db import Book, DuplicateFormulationIdError, Formulation, Sheet, SessionLocal, init_db, save_parsed_workbook
from parser import parse_workbook

st.set_page_config(page_title="試作レシピ管理", layout="wide")
init_db()

# =========================================================
# サイドバー：ログイン
#
# ローカル版（user_profile.jsonにファイルとして記憶する方式）との違い:
#   このクラウド版は複数人が同じ1つのアプリインスタンスにアクセスするため、
#   ファイルに保存すると他の人のログイン状態を上書きしてしまう。
#   そのため st.session_state（ブラウザのタブ単位のメモリ）に保持する方式にしている。
#   → ブラウザ・タブを閉じると再ログインが必要（ローカル版のような永続記憶はしない）。
# =========================================================
st.sidebar.header("ログイン")

if "profile" not in st.session_state:
    st.session_state.profile = None

if st.session_state.profile is None:
    st.sidebar.write("このタブでの利用中だけ保持されます（閉じると再入力が必要です）。")
    with st.sidebar.form("login_form"):
        email = st.text_input("メールアドレス")
        name = st.text_input("氏名")
        author_code = st.text_input("作成者コード（コード表を参照。例: tnk）")
        submitted = st.form_submit_button("ログイン")
    if submitted:
        if email and name and author_code:
            st.session_state.profile = {"email": email.strip(), "name": name.strip(), "author_code": author_code.strip()}
            st.rerun()
        else:
            st.sidebar.error("すべての項目を入力してください。")
    current_user_code = None
else:
    profile = st.session_state.profile
    st.sidebar.success(f"{profile['name']} さん")
    st.sidebar.caption(f"{profile['email']} / 作成者コード: {profile['author_code']}")
    if st.sidebar.button("ログアウト"):
        st.session_state.profile = None
        st.session_state.pop("uploaded_books", None)
        st.rerun()
    current_user_code = profile["author_code"]

st.title("試作レシピ管理（クラウド共有版）")

tab_upload, tab_mypage, tab_search = st.tabs(["アップロード", "マイページ", "検索"])

# =========================================================
# アップロードタブ
# 複数ファイルを選択・追加できるようにし、パース結果はセッション中（ログイン中）保持する。
# =========================================================
with tab_upload:
    st.write("試作ブックのExcelファイルをアップロードしてください（複数選択可）。")
    st.caption("1つのブックに複数の作成者が混在していても問題ありません。他の人が作成したブックの代理アップロードも可能です。")

    if "uploaded_books" not in st.session_state:
        st.session_state.uploaded_books = {}  # {filename: {"parsed": ParsedWorkbook, "saved": bool}}

    uploaded_files = st.file_uploader(
        "Excelファイル (.xlsx)", type=["xlsx"], accept_multiple_files=True
    )

    # 今回選択されたファイルをパースしてセッションに追加（同名ファイルは上書き＝再読み込み扱い）
    if uploaded_files:
        for uploaded in uploaded_files:
            with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
                tmp.write(uploaded.getvalue())
                tmp_path = tmp.name
            try:
                parsed = parse_workbook(tmp_path)
            except Exception as e:
                st.session_state.uploaded_books[uploaded.name] = {"parsed": None, "saved": False, "error": str(e)}
            else:
                st.session_state.uploaded_books[uploaded.name] = {"parsed": parsed, "saved": False, "error": None}

    if not st.session_state.uploaded_books:
        st.info("ファイルを選択すると、ここにプレビューが表示されます。")
    else:
        st.markdown(f"**保持中のファイル：{len(st.session_state.uploaded_books)}件**（ログイン中は保持されます）")
        for filename, entry in list(st.session_state.uploaded_books.items()):
            saved_mark = "✅ " if entry["saved"] else ""
            with st.expander(f"{saved_mark}{filename}", expanded=not entry["saved"]):
                if entry["error"]:
                    st.error(f"読み込みに失敗しました: {entry['error']}")
                    if st.button("リストから削除", key=f"remove_{filename}"):
                        del st.session_state.uploaded_books[filename]
                        st.rerun()
                    continue

                parsed = entry["parsed"]
                st.markdown("**試作ブック情報**")
                st.dataframe(pd.DataFrame([parsed.book.__dict__]), hide_index=True)

                st.markdown(f"**試作シート一覧（{len(parsed.sheets)}件）**")
                if parsed.sheets:
                    st.dataframe(pd.DataFrame([s.__dict__ for s in parsed.sheets]), hide_index=True)
                else:
                    st.info("試作シートが見つかりませんでした。")

                st.markdown(f"**配合（行9〜13）一覧（{len(parsed.formulations)}件）**")
                if parsed.formulations:
                    st.dataframe(pd.DataFrame([f.__dict__ for f in parsed.formulations]), hide_index=True)
                else:
                    st.info("配合タイトルが入力された配合が見つかりませんでした。")

                if parsed.warnings:
                    st.warning("　\n".join(["確認事項:"] + [f"- {w}" for w in parsed.warnings]))

                col_save, col_remove = st.columns([1, 1])
                with col_save:
                    if st.button("この内容をDBに保存する", type="primary", key=f"save_{filename}"):
                        try:
                            result = save_parsed_workbook(parsed, filename)
                        except DuplicateFormulationIdError as e:
                            st.error(
                                f"保存できませんでした。{e}\n\n"
                                "考えられる原因：同じ作成者コードで、ごく短時間のうちに複数のブックを新規作成した"
                                "可能性があります（ブックIDが重複しています）。"
                            )
                        else:
                            st.session_state.uploaded_books[filename]["saved"] = True
                            st.success(
                                f"保存しました（ブックID: {result['book_id']} / "
                                f"シート {result['sheets']}件 / 配合 {result['formulations']}件）"
                            )
                with col_remove:
                    if st.button("リストから削除", key=f"remove_{filename}"):
                        del st.session_state.uploaded_books[filename]
                        st.rerun()

# =========================================================
# 共通: Sheet一覧をDataFrame化するヘルパー
# book_author_code を指定すると「そのブックの作成者（表紙の作成者コード）」で絞り込む
# =========================================================
def _load_sheet_rows(book_author_code=None):
    session = SessionLocal()
    try:
        q = session.query(Sheet, Book).join(Book, Sheet.book_id == Book.id)
        if book_author_code:
            q = q.filter(Book.created_by_code == book_author_code)
        q = q.order_by(Book.book_id, Sheet.trial_no)
        rows = []
        for sheet, book in q.all():
            rows.append(
                {
                    "ブックID": book.book_id,
                    "企画名": book.project_name,
                    "カテゴリー": book.category_name or book.category_code,
                    "ブランド名": book.brand_name,
                    "試作No": sheet.trial_no,
                    "シート名": sheet.sheet_name,
                    "試作ID": sheet.formulation_id,
                    "タイトル": sheet.title,
                    "ステータス": sheet.status,
                    "作成者": sheet.created_by,
                    "作成者コード": sheet.created_by_code,
                    "試作日": sheet.trial_date,
                    "背景・目的": sheet.background_purpose,
                    "結論": sheet.conclusion,
                    "sheet_pk": sheet.id,
                }
            )
        return rows
    finally:
        session.close()


def _render_sheet_table_with_detail(rows, empty_message, key_prefix):
    if not rows:
        st.info(empty_message)
        return
    df = pd.DataFrame(rows)
    st.dataframe(df.drop(columns=["sheet_pk"]), hide_index=True, use_container_width=True)

    st.markdown("---")
    st.markdown("**配合の詳細を見る**")
    options = {f"{r['シート名']}（{r['タイトル']}・{r['作成者']}）": r["sheet_pk"] for r in rows}
    selected_label = st.selectbox("試作シートを選択", list(options.keys()), key=f"{key_prefix}_select")
    selected_sheet_id = options[selected_label]

    session = SessionLocal()
    try:
        formulations = session.query(Formulation).filter_by(sheet_id=selected_sheet_id).all()
        if formulations:
            fdf = pd.DataFrame(
                [
                    {
                        "配合No": f.formulation_no,
                        "配合タイトル": f.formulation_title,
                        "配合目的・意図": f.purpose_intent,
                        "配合結果": f.result,
                        "試作評価": f.evaluation,
                    }
                    for f in formulations
                ]
            )
            st.dataframe(fdf, hide_index=True, use_container_width=True)
        else:
            st.info("このシートには配合が登録されていません。")
    finally:
        session.close()


# =========================================================
# マイページタブ：自分が「作成者」である配合（シート単位の作成者コードで判定）
# =========================================================
with tab_mypage:
    if current_user_code is None:
        st.info("サイドバーからログインしてください。")
    else:
        st.write(f"あなた（作成者コード: {current_user_code}）が作成者になっているブックの試作一覧です。")
        st.caption("ブックの表紙にある作成者コードで判定しています。")
        rows = _load_sheet_rows(book_author_code=current_user_code)
        _render_sheet_table_with_detail(
            rows,
            empty_message="あなたが作成者になっている試作はまだ登録されていません。",
            key_prefix="mypage",
        )

# =========================================================
# 検索タブ：全員分を対象に検索・絞り込み・ソート
# =========================================================
with tab_search:
    st.write("登録済みの全ブック・全試作シートを対象に検索できます。")
    all_rows = _load_sheet_rows(book_author_code=None)

    if not all_rows:
        st.info("まだデータが登録されていません。")
    else:
        all_df = pd.DataFrame(all_rows)

        col1, col2, col3, col4 = st.columns(4)
        with col1:
            keyword = st.text_input("キーワード（タイトル・背景目的・結論）")
        with col2:
            category_options = sorted(all_df["カテゴリー"].dropna().unique().tolist())
            category_filter = st.multiselect("カテゴリー", category_options)
        with col3:
            status_options = sorted(all_df["ステータス"].dropna().unique().tolist())
            status_filter = st.multiselect("ステータス", status_options)
        with col4:
            author_options = sorted(all_df["作成者"].dropna().unique().tolist())
            author_filter = st.multiselect("作成者", author_options)

        sort_col1, sort_col2 = st.columns(2)
        with sort_col1:
            sort_key = st.selectbox(
                "並び替え", ["試作日", "タイトル", "ブックID", "試作No", "ステータス"]
            )
        with sort_col2:
            sort_desc = st.checkbox("降順にする", value=True)

        filtered = all_df.copy()
        if keyword:
            mask = (
                filtered["タイトル"].fillna("").str.contains(keyword, case=False)
                | filtered["背景・目的"].fillna("").str.contains(keyword, case=False)
                | filtered["結論"].fillna("").str.contains(keyword, case=False)
            )
            filtered = filtered[mask]
        if category_filter:
            filtered = filtered[filtered["カテゴリー"].isin(category_filter)]
        if status_filter:
            filtered = filtered[filtered["ステータス"].isin(status_filter)]
        if author_filter:
            filtered = filtered[filtered["作成者"].isin(author_filter)]

        filtered = filtered.sort_values(by=sort_key, ascending=not sort_desc, na_position="last")

        st.markdown(f"**{len(filtered)}件 / 全{len(all_df)}件**")
        rows_for_detail = filtered.to_dict("records")
        _render_sheet_table_with_detail(
            rows_for_detail,
            empty_message="条件に一致する試作が見つかりませんでした。",
            key_prefix="search",
        )
