from __future__ import annotations

import base64, pathlib
import streamlit as st
import pandas as pd
import numpy as np

from retentionsignal_core import (
    add_tpi_ranks,
    apply_tpi_formula,
    build_item_stats,
    build_student_summary,
    build_tpi_matrix,
    make_default_formula,
    parse_exam_filename,
    read_single_exam,
    to_csv_bytes,
    to_xlsx_bytes,
    ALIAS_TO_COLUMN,
    METRIC_ORDER,
)
from theme import (
    inject_custom_css,
    BLUE1, TEXT_MUTED,
)

# ── Page config ──────────────────────────────────────────────────────────────
st.set_page_config(page_title="Poly TPI Calculator", layout="wide", initial_sidebar_state="collapsed")
st.markdown(inject_custom_css(), unsafe_allow_html=True)
st.markdown("<style>section[data-testid='stSidebar']{display:none !important;}</style>", unsafe_allow_html=True)

# ── Poly logo helper ─────────────────────────────────────────────────────────
_LOGO_PNG = pathlib.Path(__file__).parent / "bi_poly.png"
_LOGO_SVG = pathlib.Path(__file__).parent / "poly_logo.svg"


def _logo_html(width: int = 260) -> str:
    for path, mime in [(_LOGO_PNG, "image/png"), (_LOGO_SVG, "image/svg+xml")]:
        if path.exists():
            b64 = base64.b64encode(path.read_bytes()).decode()
            return f'<img src="data:{mime};base64,{b64}" width="{width}" style="display:block;margin:0 auto;">'
    return f'<div style="text-align:center;font-family:Arial Black,sans-serif;font-size:72px;font-weight:900;color:{BLUE1};">Poly</div>'


LOADING_OVERLAY_HTML = """
<div class="loading-overlay">
  <div class="poly-anim-text">Poly</div>
  <div class="loading-msg">데이터 처리 중 ...</div>
</div>
"""

# ── Session state defaults ───────────────────────────────────────────────────
_DEFAULTS = {
    "data_loaded": False,
    "processing": False,
    "raw_df": None,
    "item_df": None,
    "summary_df": None,
    "exam_file_names": [],
    "tpi_result": None,
    "formula_used": None,
}
for _k, _v in _DEFAULTS.items():
    if _k not in st.session_state:
        st.session_state[_k] = _v


# ── Helpers ──────────────────────────────────────────────────────────────────
def safe_error(msg: str):
    st.markdown(f'<div class="error-box">{msg}</div>', unsafe_allow_html=True)


def load_all(exam_objs):
    exam_frames = [read_single_exam(f) for f in exam_objs]
    raw = pd.concat(exam_frames, ignore_index=True) if exam_frames else pd.DataFrame()
    item = build_item_stats(raw) if not raw.empty else pd.DataFrame()
    summary = build_student_summary(raw) if not raw.empty else pd.DataFrame()
    return raw, item, summary


# ═════════════════════════════════════════════════════════════════════════════
#  LANDING PAGE
# ═════════════════════════════════════════════════════════════════════════════
if not st.session_state.data_loaded:
    if st.session_state.processing:
        st.markdown(LOADING_OVERLAY_HTML, unsafe_allow_html=True)
        try:
            raw_df, item_df, summary_df = load_all(st.session_state._pending_exams)
            if summary_df.empty:
                safe_error("처리 가능한 데이터가 없습니다. 파일을 확인해 주세요.")
                st.session_state.processing = False
                st.stop()
            st.session_state.raw_df = raw_df
            st.session_state.item_df = item_df
            st.session_state.summary_df = summary_df
            st.session_state.data_loaded = True
            st.session_state.processing = False
            st.session_state.pop("_pending_exams", None)
            st.rerun()
        except Exception as exc:
            st.session_state.processing = False
            st.session_state.pop("_pending_exams", None)
            safe_error(f"데이터 로딩 중 오류가 발생했습니다.\n{exc}")
            st.stop()

    # ── Landing page ──
    st.markdown("<div style='height:30px'></div>", unsafe_allow_html=True)
    st.markdown(_logo_html(260), unsafe_allow_html=True)
    st.markdown('<div class="landing-title" style="text-align:center;">Poly TPI Calculator</div>', unsafe_allow_html=True)
    st.markdown('<div class="landing-subtitle" style="text-align:center;">* 문항분석표 기반 TPI 산출 *</div>', unsafe_allow_html=True)
    st.markdown("<div style='height:24px'></div>", unsafe_allow_html=True)

    _, col_center, _ = st.columns([2, 6, 2])
    with col_center:
        st.markdown(
            '<div class="upload-card"><h4>문항분석표 업로드</h4>'
            '<p class="upload-hint">시험 문항결과 .xlsx 파일 (여러 파일 가능)</p></div>',
            unsafe_allow_html=True,
        )
        exam_files = st.file_uploader(
            "문항분석표 (.xlsx)",
            type=["xlsx"],
            accept_multiple_files=True,
            key="upload_exam",
            label_visibility="collapsed",
        )

        # Show parsed info from filenames
        if exam_files:
            st.markdown("**업로드된 파일 정보:**")
            info_rows = []
            for f in exam_files:
                y, m, et, lv = parse_exam_filename(f.name)
                info_rows.append({
                    "파일명": f.name,
                    "연도": y or "-",
                    "시험유형": et or "-",
                    "월": f"{m}월" if m else "-",
                    "레벨": lv or "-",
                })
            st.dataframe(pd.DataFrame(info_rows), use_container_width=True, hide_index=True)

    st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)
    _, btn_col, _ = st.columns([4, 3, 4])
    with btn_col:
        analyze_clicked = st.button("데이터 로드", type="primary", use_container_width=True)

    if analyze_clicked:
        if not exam_files:
            safe_error("문항분석표 파일을 업로드해 주세요.")
            st.stop()
        st.session_state.exam_file_names = [f.name for f in exam_files]
        st.session_state._pending_exams = exam_files
        st.session_state.processing = True
        st.rerun()

    st.stop()


# ═════════════════════════════════════════════════════════════════════════════
#  MAIN APPLICATION (data loaded)
# ═════════════════════════════════════════════════════════════════════════════

raw_df = st.session_state.raw_df
item_df = st.session_state.item_df
summary_df = st.session_state.summary_df

# ── File accordion (top) ────────────────────────────────────────────────────
exam_names = st.session_state.exam_file_names
file_count = len(exam_names)

accordion_html = f"""
<details class="file-accordion">
  <summary>업로드된 파일 ({file_count}개)</summary>
  <div class="file-list">
"""
for fn in exam_names:
    accordion_html += f"📊 {fn}<br>"
accordion_html += "</div></details>"

col_acc, col_reset = st.columns([9, 1])
with col_acc:
    st.markdown(accordion_html, unsafe_allow_html=True)
with col_reset:
    if st.button("새 분석", key="btn_reset"):
        for k in list(st.session_state.keys()):
            del st.session_state[k]
        st.rerun()

# ── Filter controls ──────────────────────────────────────────────────────────
all_exam_types = sorted(summary_df["시험유형"].dropna().unique().tolist())
all_month_opts = sorted(summary_df["월"].dropna().astype(int).unique().tolist())
has_level = "레벨" in summary_df.columns
all_levels = sorted(summary_df["레벨"].dropna().unique().tolist()) if has_level else []
campus_opts = sorted(summary_df["캠퍼스"].dropna().unique().tolist()) if "캠퍼스" in summary_df.columns else []

st.markdown('<div class="filter-container">', unsafe_allow_html=True)
st.markdown('<div class="filter-header">🔍 조회 조건</div>', unsafe_allow_html=True)

n_filter_cols = 3 + (1 if has_level else 0) + (1 if campus_opts else 0)
filter_cols = st.columns(n_filter_cols)

col_idx = 0
with filter_cols[col_idx]:
    st.markdown('<div class="filter-label">📋 시험유형</div>', unsafe_allow_html=True)
    selected_exam_types = st.multiselect(
        "시험유형", options=all_exam_types, default=all_exam_types,
        key="filter_exam_type", label_visibility="collapsed",
    )
col_idx += 1

with filter_cols[col_idx]:
    st.markdown('<div class="filter-label">📅 월</div>', unsafe_allow_html=True)
    selected_months = st.multiselect(
        "월", options=all_month_opts, default=all_month_opts,
        key="filter_month", label_visibility="collapsed",
    )
col_idx += 1

if has_level:
    with filter_cols[col_idx]:
        st.markdown('<div class="filter-label">📚 레벨</div>', unsafe_allow_html=True)
        selected_levels = st.multiselect(
            "레벨", options=all_levels, default=all_levels,
            key="filter_level", label_visibility="collapsed",
        )
    col_idx += 1
else:
    selected_levels = []

if campus_opts:
    with filter_cols[col_idx]:
        st.markdown('<div class="filter-label">🏫 캠퍼스</div>', unsafe_allow_html=True)
        selected_campuses = st.multiselect(
            "캠퍼스", options=campus_opts, default=campus_opts,
            key="filter_campus", label_visibility="collapsed",
        )
    col_idx += 1
else:
    selected_campuses = []

with filter_cols[col_idx]:
    st.markdown('<div class="filter-label">👤 학생코드 (선택)</div>', unsafe_allow_html=True)
    selected_students = st.multiselect(
        "학생코드", options=summary_df["학생코드"].dropna().astype(str).unique().tolist(),
        default=[], key="filter_student", label_visibility="collapsed",
    )

st.markdown('</div>', unsafe_allow_html=True)

# Apply filters
view_df = summary_df[
    summary_df["시험유형"].isin(selected_exam_types) & summary_df["월"].isin(selected_months)
].copy()
if has_level and selected_levels:
    view_df = view_df[view_df["레벨"].isin(selected_levels)].copy()
if selected_campuses and "캠퍼스" in view_df.columns:
    view_df = view_df[view_df["캠퍼스"].isin(selected_campuses)].copy()
if selected_students:
    view_df = view_df[view_df["학생코드"].astype(str).isin(selected_students)].copy()

# ── Main Tabs ────────────────────────────────────────────────────────────────
tab1, tab2, tab3 = st.tabs(["📊 TPI 계산", "📋 학생 성적표", "📝 문항 정답률"])

# ═══ TAB 1: TPI 계산 ═══
with tab1:
    st.subheader("TPI 비율 설정")
    st.markdown("사용 가능한 별칭: `P`, `T`, `BCV`, `CI`, `QR`, `CT`, `CCV`, `CQR`, `CV`  ·  연산자: `+`, `-`, `*`, `/`, `**`, `( )`")

    formula_mode = st.radio(
        "수식 입력 방식", ["가중치(비율) 기반 자동 생성", "자유 수식 직접 입력"], horizontal=True, key="formula_mode"
    )

    if formula_mode == "가중치(비율) 기반 자동 생성":
        default_weights = {"P": 20, "T": 20, "BCV": 15, "CI": 15, "QR": 10, "CT": 10, "CCV": 5, "CQR": 5, "CV": 0}
        enabled_weights = {}
        wcols = st.columns(5)
        aliases = list(default_weights.keys())
        for i, alias in enumerate(aliases):
            with wcols[i % 5]:
                default_on = alias != "CV"
                enabled = st.checkbox(f"{alias} 채택", value=default_on, key=f"en_{alias}")
                weight = st.number_input(
                    f"{alias} 비율%", min_value=0.0, max_value=100.0,
                    value=float(default_weights[alias]), step=1.0, key=f"wt_{alias}",
                )
                enabled_weights[alias] = weight if enabled else 0.0
        formula = make_default_formula(enabled_weights)
        st.text_input("생성된 TPI 수식", value=formula, disabled=True, key="formula_display")
    else:
        formula = st.text_input(
            "TPI 수식",
            value="(P*20.0 + T*20.0 + BCV*15.0 + CI*15.0 + QR*10.0 + CT*10.0 + CCV*5.0 + CQR*5.0) / 100.0",
            key="formula_free",
        )

    # Calculate button
    tpi_run = st.button("계 산", type="primary", key="btn_tpi_run")

    if tpi_run:
        with st.spinner("TPI 계산 중..."):
            try:
                tpi_df = apply_tpi_formula(view_df, formula)
                tpi_df = add_tpi_ranks(tpi_df)
                st.session_state.tpi_result = tpi_df
                st.session_state.formula_used = formula
                st.success("TPI 계산 완료")
            except Exception as e:
                safe_error(f"TPI 수식 오류: {e}")

    # Show TPI result table
    if st.session_state.tpi_result is not None:
        tpi_df = st.session_state.tpi_result

        # Build display columns
        display_cols = ["캠퍼스", "학생코드", "학생명", "시험유형", "월"]
        if has_level:
            display_cols.append("레벨")
        display_cols += ["TPI", "TPI랭크(전체)", "TPI랭크(캠퍼스)"]
        display_cols += ["P-Score", "T-Score", "B.CV", "CI", "QR", "C.T-Score", "C.CV", "C.QR"]

        # Filter to only columns that exist
        display_cols = [c for c in display_cols if c in tpi_df.columns]
        result_view = tpi_df[display_cols].copy()

        st.dataframe(result_view, use_container_width=True, height=650)

        # Download buttons
        dl_col1, dl_col2 = st.columns(2)
        with dl_col1:
            st.download_button(
                "📥 XLSX 다운로드",
                data=to_xlsx_bytes(result_view),
                file_name="Poly_TPI_Result.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="btn_dl_xlsx",
            )
        with dl_col2:
            st.download_button(
                "📥 CSV 다운로드",
                data=to_csv_bytes(result_view),
                file_name="Poly_TPI_Result.csv",
                mime="text/csv",
                key="btn_dl_csv",
            )

    with st.expander("TPI 지표 별칭(Alias) 가이드"):
        st.markdown("""
- **P**: P-Score (전체 정답률)  · **T**: T-Score (표준점수 기반 0-100)
- **BCV**: B.CV (과목별 편차 역수)  · **CI**: CI (난이도 가중 성취 지표)
- **QR**: QR (백분위 순위)  · **CT**: C.T-Score (캠퍼스 내 T-Score)
- **CCV**: C.CV (캠퍼스 내 변동성)  · **CQR**: C.QR (캠퍼스 내 백분위)  · **CV**: CV (전체 시험 변동성)
""")

# ═══ TAB 2: 학생 성적표 ═══
with tab2:
    st.subheader("학생 성적표")
    st.dataframe(view_df, use_container_width=True, height=600)
    dl1, dl2 = st.columns(2)
    with dl1:
        st.download_button(
            "📥 XLSX 다운로드",
            data=to_xlsx_bytes(view_df),
            file_name="Poly_Student_Summary.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="btn_dl_student_xlsx",
        )
    with dl2:
        st.download_button(
            "📥 CSV 다운로드",
            data=to_csv_bytes(view_df),
            file_name="Poly_Student_Summary.csv",
            mime="text/csv",
            key="btn_dl_student_csv",
        )

# ═══ TAB 3: 문항 정답률 ═══
with tab3:
    st.subheader("문항 정답률")
    item_view = item_df[
        item_df["exam_type"].isin(selected_exam_types) & item_df["month_num"].isin(selected_months)
    ].copy()
    if has_level and "level" in item_df.columns and selected_levels:
        item_view = item_view[item_view["level"].isin(selected_levels)].copy()
    item_view = item_view.rename(columns={"exam_type": "시험유형", "year": "연도", "month_num": "월"})
    if "level" in item_view.columns:
        item_view = item_view.rename(columns={"level": "레벨"})
    st.dataframe(item_view, use_container_width=True, height=600)
    dl1, dl2 = st.columns(2)
    with dl1:
        st.download_button(
            "📥 XLSX 다운로드",
            data=to_xlsx_bytes(item_view),
            file_name="Poly_Item_Accuracy.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="btn_dl_item_xlsx",
        )
    with dl2:
        st.download_button(
            "📥 CSV 다운로드",
            data=to_csv_bytes(item_view),
            file_name="Poly_Item_Accuracy.csv",
            mime="text/csv",
            key="btn_dl_item_csv",
        )
