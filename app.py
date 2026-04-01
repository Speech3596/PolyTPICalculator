from __future__ import annotations

import base64, pathlib
import streamlit as st
import pandas as pd
import numpy as np

from retentionsignal_core import (
    apply_tpi_formula,
    build_item_stats,
    build_student_summary,
    compute_risk_grades,
    make_default_formula,
    parse_exam_filename,
    read_single_exam,
    to_csv_bytes,
    to_xlsx_bytes,
    ALIAS_TO_COLUMN,
    METRIC_ORDER,
    RISK_GRADE_DESCRIPTIONS,
)
from theme import inject_custom_css, BLUE1, TEXT_MUTED

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(page_title="Poly TPI Calculator", layout="wide", initial_sidebar_state="collapsed")
st.markdown(inject_custom_css(), unsafe_allow_html=True)
st.markdown(
    "<style>section[data-testid='stSidebar']{display:none !important;}</style>",
    unsafe_allow_html=True,
)

# ── Logo ─────────────────────────────────────────────────────────────────────
_LOGO_PNG = pathlib.Path(__file__).parent / "bi_poly.png"
_LOGO_SVG = pathlib.Path(__file__).parent / "poly_logo.svg"


def _logo_html(width: int = 260) -> str:
    for path, mime in [(_LOGO_PNG, "image/png"), (_LOGO_SVG, "image/svg+xml")]:
        if path.exists():
            b64 = base64.b64encode(path.read_bytes()).decode()
            return f'<img src="data:{mime};base64,{b64}" width="{width}" style="display:block;margin:0 auto;">'
    return (
        f'<div style="text-align:center;font-family:Arial Black,sans-serif;'
        f'font-size:72px;font-weight:900;color:{BLUE1};">Poly</div>'
    )


LOADING_OVERLAY_HTML = """
<div class="loading-overlay">
  <div class="poly-anim-text">Poly</div>
  <div class="loading-msg">데이터 처리 중 ...</div>
</div>
"""

# ── Session state ─────────────────────────────────────────────────────────────
_DEFAULTS = {
    "data_loaded":     False,
    "processing":      False,
    "raw_df":          None,
    "item_df":         None,
    "summary_df":      None,
    "exam_file_names": [],
    "tpi_result":      None,   # enriched df after compute_risk_grades
    "formula_used":    None,
}
for _k, _v in _DEFAULTS.items():
    if _k not in st.session_state:
        st.session_state[_k] = _v


# ── Helpers ───────────────────────────────────────────────────────────────────
def safe_error(msg: str):
    st.markdown(f'<div class="error-box">{msg}</div>', unsafe_allow_html=True)


def load_all(exam_objs):
    from retentionsignal_core import read_single_exam, build_item_stats, build_student_summary
    frames = [read_single_exam(f) for f in exam_objs]
    raw = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
    item = build_item_stats(raw) if not raw.empty else pd.DataFrame()
    summary = build_student_summary(raw) if not raw.empty else pd.DataFrame()
    return raw, item, summary


# ─────────────────────────────────────────────────────────────────────────────
#  LANDING PAGE
# ─────────────────────────────────────────────────────────────────────────────
if not st.session_state.data_loaded:
    if st.session_state.processing:
        st.markdown(LOADING_OVERLAY_HTML, unsafe_allow_html=True)
        try:
            raw_df, item_df, summary_df = load_all(st.session_state._pending_exams)
            if summary_df.empty:
                safe_error("처리 가능한 데이터가 없습니다. 파일을 확인해 주세요.")
                st.session_state.processing = False
                st.stop()
            # Auto-compute TPI with default weights immediately after loading
            _default_formula = (
                "(T*30.0 + TENG*10.0 + TENGF*5.0 + TSB*5.0 + QR*20.0 + BCV*30.0) / 100.0"
            )
            try:
                _full_tpi = apply_tpi_formula(summary_df, _default_formula)
                _full_enriched = compute_risk_grades(_full_tpi)
                st.session_state.tpi_result   = _full_enriched
                st.session_state.formula_used = _default_formula
            except Exception:
                st.session_state.tpi_result   = None
                st.session_state.formula_used = None
            st.session_state.raw_df    = raw_df
            st.session_state.item_df   = item_df
            st.session_state.summary_df = summary_df
            st.session_state.data_loaded = True
            st.session_state.processing  = False
            st.session_state.pop("_pending_exams", None)
            st.rerun()
        except Exception as exc:
            st.session_state.processing = False
            st.session_state.pop("_pending_exams", None)
            safe_error(f"데이터 로딩 오류: {exc}")
            st.stop()

    # Landing UI
    st.markdown("<div style='height:30px'></div>", unsafe_allow_html=True)
    st.markdown(_logo_html(260), unsafe_allow_html=True)
    st.markdown('<div class="landing-title" style="text-align:center;">Poly TPI Calculator</div>', unsafe_allow_html=True)
    st.markdown('<div class="landing-subtitle" style="text-align:center;">* 문항분석표 기반 TPI 산출 *</div>', unsafe_allow_html=True)
    st.markdown("<div style='height:24px'></div>", unsafe_allow_html=True)

    _, col_c, _ = st.columns([2, 6, 2])
    with col_c:
        st.markdown(
            '<div class="upload-card"><h4>문항분석표 업로드</h4>'
            '<p class="upload-hint">시험 문항결과 .xlsx 파일 — 여러 파일 동시 업로드 가능</p></div>',
            unsafe_allow_html=True,
        )
        exam_files = st.file_uploader(
            "문항분석표 (.xlsx)",
            type=["xlsx"],
            accept_multiple_files=True,
            key="upload_exam",
            label_visibility="collapsed",
        )
        if exam_files:
            rows = []
            for f in exam_files:
                y, m, et, lv = parse_exam_filename(f.name)
                rows.append({
                    "파일명": f.name,
                    "연도": y or "-",
                    "시험유형": et or "-",
                    "월": f"{m}월" if m else "-",
                    "레벨": lv or "-",
                })
            st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

    st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)
    _, btn_col, _ = st.columns([4, 3, 4])
    with btn_col:
        if st.button("데이터 로드", type="primary", use_container_width=True):
            if not exam_files:
                safe_error("문항분석표 파일을 업로드해 주세요.")
                st.stop()
            st.session_state.exam_file_names = [f.name for f in exam_files]
            st.session_state._pending_exams  = exam_files
            st.session_state.processing      = True
            st.rerun()

    st.stop()


# ─────────────────────────────────────────────────────────────────────────────
#  MAIN APPLICATION
# ─────────────────────────────────────────────────────────────────────────────
raw_df     = st.session_state.raw_df
item_df    = st.session_state.item_df
summary_df = st.session_state.summary_df

# ── File accordion + reset ────────────────────────────────────────────────────
exam_names = st.session_state.exam_file_names
html_acc = (
    f'<details class="file-accordion"><summary>업로드된 파일 ({len(exam_names)}개)</summary>'
    f'<div class="file-list">' +
    "".join(f"📊 {fn}<br>" for fn in exam_names) +
    "</div></details>"
)
col_acc, col_rst = st.columns([9, 1])
with col_acc:
    st.markdown(html_acc, unsafe_allow_html=True)
with col_rst:
    if st.button("새 분석", key="btn_reset"):
        for k in list(st.session_state.keys()):
            del st.session_state[k]
        st.rerun()

# ── Filter bar ────────────────────────────────────────────────────────────────
all_exam_types = sorted(summary_df["시험유형"].dropna().unique().tolist())
all_months     = sorted(summary_df["월"].dropna().astype(int).unique().tolist())
has_level      = "레벨" in summary_df.columns
all_levels     = sorted(summary_df["레벨"].dropna().unique().tolist()) if has_level else []
campus_opts    = sorted(summary_df["캠퍼스"].dropna().unique().tolist()) if "캠퍼스" in summary_df.columns else []

st.markdown('<div class="filter-container">', unsafe_allow_html=True)
st.markdown('<div class="filter-header">🔍 조회 조건</div>', unsafe_allow_html=True)

_n_cols = 2 + (1 if has_level else 0) + (1 if campus_opts else 0) + 1
_fcols  = st.columns(_n_cols)
_ci = 0

with _fcols[_ci]:
    st.markdown('<div class="filter-label">📋 시험유형</div>', unsafe_allow_html=True)
    selected_exam_types = st.multiselect(
        "시험유형", options=all_exam_types, default=all_exam_types,
        key="f_et", label_visibility="collapsed",
    )
_ci += 1

with _fcols[_ci]:
    st.markdown('<div class="filter-label">📅 월</div>', unsafe_allow_html=True)
    selected_months = st.multiselect(
        "월", options=all_months, default=all_months,
        key="f_mo", label_visibility="collapsed",
    )
_ci += 1

if has_level:
    with _fcols[_ci]:
        st.markdown('<div class="filter-label">📚 레벨</div>', unsafe_allow_html=True)
        selected_levels = st.multiselect(
            "레벨", options=all_levels, default=all_levels,
            key="f_lv", label_visibility="collapsed",
        )
    _ci += 1
else:
    selected_levels = []

if campus_opts:
    with _fcols[_ci]:
        st.markdown('<div class="filter-label">🏫 캠퍼스</div>', unsafe_allow_html=True)
        selected_campuses = st.multiselect(
            "캠퍼스", options=campus_opts, default=campus_opts,
            key="f_ca", label_visibility="collapsed",
        )
    _ci += 1
else:
    selected_campuses = []

with _fcols[_ci]:
    st.markdown('<div class="filter-label">👤 학생코드 (선택)</div>', unsafe_allow_html=True)
    selected_students = st.multiselect(
        "학생코드",
        options=summary_df["학생코드"].dropna().astype(str).unique().tolist(),
        default=[],
        key="f_stu",
        label_visibility="collapsed",
    )

st.markdown("</div>", unsafe_allow_html=True)

# Apply filters to summary_df
view_df = summary_df[
    summary_df["시험유형"].isin(selected_exam_types) & summary_df["월"].isin(selected_months)
].copy()
if has_level and selected_levels:
    view_df = view_df[view_df["레벨"].isin(selected_levels)].copy()
if selected_campuses and "캠퍼스" in view_df.columns:
    view_df = view_df[view_df["캠퍼스"].isin(selected_campuses)].copy()
if selected_students:
    view_df = view_df[view_df["학생코드"].astype(str).isin(selected_students)].copy()

# Apply same filters to tpi_result if present
tpi_result = st.session_state.tpi_result
if tpi_result is not None:
    tpi_view = tpi_result[
        tpi_result["시험유형"].isin(selected_exam_types) & tpi_result["월"].isin(selected_months)
    ].copy()
    if has_level and selected_levels and "레벨" in tpi_view.columns:
        tpi_view = tpi_view[tpi_view["레벨"].isin(selected_levels)].copy()
    if selected_campuses and "캠퍼스" in tpi_view.columns:
        tpi_view = tpi_view[tpi_view["캠퍼스"].isin(selected_campuses)].copy()
    if selected_students and "학생코드" in tpi_view.columns:
        tpi_view = tpi_view[tpi_view["학생코드"].astype(str).isin(selected_students)].copy()
else:
    tpi_view = None

# ─────────────────────────────────────────────────────────────────────────────
#  TABS
# ─────────────────────────────────────────────────────────────────────────────
tab_tpi, tab_score, tab_item, tab_desc = st.tabs([
    "📊 TPI 계산",
    "📋 학생 성적표",
    "📝 문항 정답률",
    "ℹ️ 지표 설명",
])

# ══════════════════════════════════════════════════════════════════════════════
#  TAB 1 — TPI 계산
# ══════════════════════════════════════════════════════════════════════════════
with tab_tpi:
    st.subheader("TPI 비율 설정")
    st.markdown(
        "별칭: `T` `TENG` `TENGF` `TSB` `QR` `BCV` `P`  ·  연산자: `+` `-` `*` `/` `**` `( )`"
    )

    formula_mode = st.radio(
        "수식 입력 방식",
        ["가중치(비율) 기반 자동 생성", "자유 수식 직접 입력"],
        horizontal=True,
        key="formula_mode",
    )

    # Default weights (T 30, TENG 10, TENGF 5, TSB 5, QR 20, BCV 30, P 0/off)
    _default_w: dict[str, tuple[float, bool]] = {
        "T":    (30.0, True),
        "TENG": (10.0, True),
        "TENGF":(5.0,  True),
        "TSB":  (5.0,  True),
        "QR":   (20.0, True),
        "BCV":  (30.0, True),
        "P":    (0.0,  False),   # disabled by default
    }
    _alias_label = {
        "T":    "T (T-Score 총점)",
        "TENG": "TENG (T-Eng)",
        "TENGF":"TENGF (T-Eng.F)",
        "TSB":  "TSB (T-S.B)",
        "QR":   "QR (백분위순위)",
        "BCV":  "BCV (B.CV)",
        "P":    "P (P-Score)",
    }

    if formula_mode == "가중치(비율) 기반 자동 생성":
        enabled_weights: dict[str, float] = {}
        cols_w = st.columns(len(_default_w))
        for i, (alias, (default_val, default_on)) in enumerate(_default_w.items()):
            with cols_w[i]:
                on = st.checkbox(
                    _alias_label[alias],
                    value=default_on,
                    key=f"en_{alias}",
                )
                wt = st.number_input(
                    f"{alias} %",
                    min_value=0.0, max_value=100.0,
                    value=default_val if default_on else 0.0,
                    step=1.0,
                    key=f"wt_{alias}",
                    label_visibility="collapsed",
                )
                enabled_weights[alias] = wt if on else 0.0

        total_w = sum(enabled_weights.values())
        formula = make_default_formula(enabled_weights)
        st.markdown(
            f"**합계: {total_w:.0f}%** {'✅' if abs(total_w - 100) < 0.01 else '⚠️ 합계가 100%가 아닙니다'}"
        )
        st.text_input("생성된 TPI 수식", value=formula, disabled=True, key="formula_display")
    else:
        formula = st.text_input(
            "TPI 수식",
            value="(T*30.0 + TENG*10.0 + TENGF*5.0 + TSB*5.0 + QR*20.0 + BCV*30.0) / 100.0",
            key="formula_free",
        )

    # ── Calculate button ──────────────────────────────────────────────────────
    tpi_run = st.button("계 산", type="primary", key="btn_tpi_run", use_container_width=False)

    if tpi_run:
        with st.spinner("TPI 계산 및 위험등급 산출 중..."):
            try:
                # Apply TPI to ALL summary rows (for correct percentile calculation)
                full_tpi = apply_tpi_formula(summary_df, formula)
                full_enriched = compute_risk_grades(full_tpi)
                st.session_state.tpi_result  = full_enriched
                st.session_state.formula_used = formula
                st.success("TPI 계산 완료")
                # Recompute tpi_view with fresh data
                tpi_view = full_enriched[
                    full_enriched["시험유형"].isin(selected_exam_types)
                    & full_enriched["월"].isin(selected_months)
                ].copy()
                if has_level and selected_levels and "레벨" in tpi_view.columns:
                    tpi_view = tpi_view[tpi_view["레벨"].isin(selected_levels)].copy()
                if selected_campuses and "캠퍼스" in tpi_view.columns:
                    tpi_view = tpi_view[tpi_view["캠퍼스"].isin(selected_campuses)].copy()
                if selected_students and "학생코드" in tpi_view.columns:
                    tpi_view = tpi_view[tpi_view["학생코드"].astype(str).isin(selected_students)].copy()
            except Exception as e:
                safe_error(f"TPI 수식 오류: {e}")

    # ── TPI Result table ──────────────────────────────────────────────────────
    if tpi_view is not None:
        st.markdown("---")
        st.markdown("### TPI 계산 결과")

        # Column order for TPI tab
        _tpi_cols = ["캠퍼스", "학생코드", "학생명", "시험유형", "월"]
        if has_level:
            _tpi_cols.append("레벨")
        _tpi_cols += [
            "TPI", "TPI랭크(전체)", "TPI랭크(캠퍼스)",
            "P-Score", "T-Score", "T-Eng", "T-Eng.F", "T-S.B", "QR", "B.CV",
        ]
        _tpi_cols = [c for c in _tpi_cols if c in tpi_view.columns]
        result_tpi = tpi_view[_tpi_cols].copy()

        st.dataframe(result_tpi, use_container_width=True, height=620)

        dl1, dl2 = st.columns(2)
        with dl1:
            st.download_button(
                "📥 XLSX 다운로드",
                data=to_xlsx_bytes(result_tpi, "TPI결과"),
                file_name="Poly_TPI_Result.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_tpi_xlsx",
            )
        with dl2:
            st.download_button(
                "📥 CSV 다운로드",
                data=to_csv_bytes(result_tpi),
                file_name="Poly_TPI_Result.csv",
                mime="text/csv",
                key="dl_tpi_csv",
            )

    with st.expander("TPI 별칭(Alias) 가이드"):
        st.markdown("""
| 별칭 | 컬럼 | 설명 |
|------|------|------|
| `T` | T-Score | 전체 P-Score 기반 표준점수 (평균 50, SD 10, 범위 0-100) |
| `TENG` | T-Eng | English 과목 T-Score |
| `TENGF` | T-Eng.F | Eng. Foundations 과목 T-Score |
| `TSB` | T-S.B | Speech Building 과목 T-Score |
| `QR` | QR | P-Score 백분위 순위 (1위=100%) |
| `BCV` | B.CV | 3개 과목 점수 편차 역수 (100-CV, 높을수록 균형) |
| `P` | P-Score | 전체 정답률 (기본 비활성) |
""")


# ══════════════════════════════════════════════════════════════════════════════
#  TAB 2 — 학생 성적표
# ══════════════════════════════════════════════════════════════════════════════
with tab_score:
    st.subheader("학생 성적표")

    # Use TPI-enriched data if available, else raw summary
    if tpi_view is not None:
        display_df = tpi_view.copy()
    else:
        display_df = view_df.copy()
        st.info("TPI 계산 탭에서 '계 산'을 클릭하면 TPI · 위험등급 · 사유 컬럼도 표시됩니다.")

    # Build ordered column list for 학생 성적표
    # Base ID columns
    _sc_id = ["캠퍼스", "학생코드", "학생명"]
    if has_level and "레벨" in display_df.columns:
        _sc_id.append("레벨")
    _sc_id += ["시험유형", "연도", "월"]

    # Subject raw score columns (sorted, but Speech Building last of the group for ordering)
    _all_cols = display_df.columns.tolist()
    _known = set(
        _sc_id
        + ["P-Score", "T-Score", "T-Eng", "T-Eng.F", "T-S.B", "QR", "B.CV",
           "TPI", "TPI랭크(전체)", "TPI랭크(캠퍼스)", "TPI분위",
           "위험등급", "사유"]
    )
    _subj_cols = sorted([c for c in _all_cols if c not in _known])

    _sc_ordered = (
        _sc_id
        + ["P-Score"]
        + _subj_cols
        + ["T-Score", "T-Eng", "T-Eng.F", "T-S.B"]
        + ["QR", "B.CV"]
    )
    if tpi_view is not None:
        _sc_ordered += ["TPI", "TPI랭크(전체)", "TPI랭크(캠퍼스)", "TPI분위", "위험등급", "사유"]

    _sc_ordered = [c for c in _sc_ordered if c in display_df.columns]
    display_df = display_df[_sc_ordered]

    # Colour-code 위험등급 column if present
    def _risk_color(val):
        colours = {
            "At-Risk":       "background-color:#FF4B4B;color:white;font-weight:bold",
            "High-Risk":     "background-color:#FF8C00;color:white;font-weight:bold",
            "Latent Risk":   "background-color:#FFA500;color:white;font-weight:bold",
            "Local Risk":    "background-color:#FFD700;color:#333;font-weight:bold",
            "Top Risk":      "background-color:#1F4E79;color:white;font-weight:bold",
            "Local Top Risk":"background-color:#2E86AB;color:white;font-weight:bold",
        }
        return colours.get(val, "")

    if "위험등급" in display_df.columns:
        try:
            # pandas >= 2.1 uses .map(); older versions use .applymap()
            styled = display_df.style.map(_risk_color, subset=["위험등급"])
        except AttributeError:
            styled = display_df.style.applymap(_risk_color, subset=["위험등급"])
        st.dataframe(styled, use_container_width=True, height=650)
    else:
        st.dataframe(display_df, use_container_width=True, height=650)

    dl1, dl2 = st.columns(2)
    with dl1:
        st.download_button(
            "📥 XLSX 다운로드",
            data=to_xlsx_bytes(display_df, "학생성적표"),
            file_name="Poly_Student_Summary.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl_score_xlsx",
        )
    with dl2:
        st.download_button(
            "📥 CSV 다운로드",
            data=to_csv_bytes(display_df),
            file_name="Poly_Student_Summary.csv",
            mime="text/csv",
            key="dl_score_csv",
        )

    # ── Risk grade summary (if TPI calculated) ─────────────────────────────
    if tpi_view is not None and "위험등급" in tpi_view.columns:
        with st.expander("위험등급 요약"):
            _rc = (
                tpi_view["위험등급"]
                .replace("", "정상")
                .fillna("정상")
                .value_counts()
                .reset_index()
                .rename(columns={"위험등급": "위험등급", "count": "학생 수"})
            )
            st.dataframe(_rc, use_container_width=True, hide_index=True)


# ══════════════════════════════════════════════════════════════════════════════
#  TAB 3 — 문항 정답률
# ══════════════════════════════════════════════════════════════════════════════
with tab_item:
    st.subheader("문항 정답률")
    item_view = item_df[
        item_df["exam_type"].isin(selected_exam_types)
        & item_df["month_num"].isin(selected_months)
    ].copy()
    if has_level and "level" in item_view.columns and selected_levels:
        item_view = item_view[item_view["level"].isin(selected_levels)].copy()
    rename_map = {"exam_type": "시험유형", "year": "연도", "month_num": "월", "level": "레벨"}
    item_view = item_view.rename(columns={k: v for k, v in rename_map.items() if k in item_view.columns})

    st.dataframe(item_view, use_container_width=True, height=600)
    dl1, dl2 = st.columns(2)
    with dl1:
        st.download_button(
            "📥 XLSX 다운로드",
            data=to_xlsx_bytes(item_view, "문항정답률"),
            file_name="Poly_Item_Accuracy.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl_item_xlsx",
        )
    with dl2:
        st.download_button(
            "📥 CSV 다운로드",
            data=to_csv_bytes(item_view),
            file_name="Poly_Item_Accuracy.csv",
            mime="text/csv",
            key="dl_item_csv",
        )


# ══════════════════════════════════════════════════════════════════════════════
#  TAB 4 — 지표 설명
# ══════════════════════════════════════════════════════════════════════════════
with tab_desc:
    st.subheader("지표 설명")

    # ── TPI 지수 평가 구조 개요 ────────────────────────────────────────────────
    st.markdown("## 📌 TPI 지수 평가 구조")
    st.markdown("""
**종합 성취도 (70%)** + **과목 간 편차 B.CV (30%)** = 총 100% 산출

---

### 성취도 (70%) 세부 분할 및 채택 근거

**전체 T-Score — 난이도 통제 및 성장 추적**
시험 난이도에 따른 원점수 왜곡을 배제하고, 절대 기준(평균 50)을 통해 과거 대비 실질적 성취 증감을 객관적으로 추적합니다.

**QR 백분위 — 수준별 정밀 측정 및 T-Score 보완**
원점수(P-Score) 기반 지표로 전체 실력을 가늠하고 T-Score의 맹점을 보완합니다.
학생에게 등수를 미공개함에도 시스템에 포함한 이유는, 동일 점수대나 상위권 내에 밀집한 학생들의 밀집도와 '실제 수준 차이'에 따른 상대적 효능감이 나타나기 때문입니다.
AUC(예측력)에서 P-Score·T-Score와 유사하거나 동일하면서도, T-Score로 가늠하기 힘든 상대적 위치 정보를 더 명확히 나타냅니다.

**과목별 T-Score — 실질 영향도 기반 정밀 가중치**
단순 총점 합산 시 발생하는 특정 과목 결손 착시를 방지합니다.
실제 데이터 분석 결과 확인된 종합 성취도 기여도 (Reading > Grammar > Vocabulary)를 근거로,
T.English 10% · T.Eng F 5% · T.S.B 5%의 가중치를 차등 배정하여 과목별 실제 중요도를 추가 반영합니다.
""")

    st.markdown("---")
    st.markdown("## 📐 TPI 구성 지표 상세")

    st.markdown("""
### P-Score (전체 정답률)
- **정의**: 전체 문항 중 정답을 맞춘 비율 × 100
- **공식**: `P-Score = (총 정답 수 / 총 문항 수) × 100`
- **범위**: 0 ~ 100
- **참고**: TPI 기본 구성에서는 비활성(체크 해제) 상태. 선택적으로 추가 가능.

---

### T-Score (총점 T점수)
- **정의**: 같은 시험·레벨 내에서 P-Score를 표준화한 점수
- **공식**: `T-Score = 50 + 10 × ((P-Score − μ) / σ)`, 범위 클리핑 [0, 100]
  - μ = 같은 (시험유형 · 월 · 레벨) 내 P-Score 평균
  - σ = 표준편차 (ddof=0)
- **해석**: 50 = 평균, 60 = 평균+1SD, 40 = 평균-1SD

---

### T-Eng / T-Eng.F / T-S.B (과목별 T점수)
- **정의**: English · Eng. Foundations · Speech Building 각 과목 점수의 T-Score 환산
- **공식**: 각 과목 점수에 대해 같은 (시험유형 · 월 · 레벨) 내 T-Score 산출
  - `T-Eng  = tscore(English 과목 점수)`
  - `T-Eng.F = tscore(Eng. Foundations 과목 점수)`
  - `T-S.B  = tscore(Speech Building 과목 점수)`
- **주의**: 다른 과목(NF Studies 등)은 T점수를 산출하지 않음

---

### QR (Quantile Rank — 백분위 순위)
- **정의**: P-Score 기준 같은 (시험유형 · 월 · 레벨) 내 백분위 순위
- **공식**: `QR = (P-Score의 백분위) × 100`
- **해석**: 1등 학생 = 100, 꼴찌 학생 ≈ 0

---

### B.CV (과목 간 편차 역수)
- **정의**: English · Eng. Foundations · Speech Building 3개 과목 점수 간의 균형도
- **공식**: `B.CV = 100 − (σ_subjects / μ_subjects × 100)`, 클리핑 [0, 100]
  - σ = 3개 과목 점수의 표준편차 (ddof=0)
  - μ = 3개 과목 점수의 평균
- **해석**: 100 = 3과목 점수가 완전히 균일, 낮을수록 과목 간 점수 편차가 큼
- **주의**: NF Studies 등 다른 과목은 B.CV 계산에서 제외

---

### TPI (Total Performance Index)
- **정의**: 사용자 지정 가중치로 구성된 종합 성취 지수
- **기본 공식**:
  ```
  TPI = (T × 30 + TENG × 10 + TENGF × 5 + TSB × 5 + QR × 20 + BCV × 30) / 100
  ```
- **범위**: 0 ~ 100

---

### TPI분위 (TPI 백분위)
- **정의**: 같은 (시험유형 · 월 · 레벨) 내 TPI 백분위 순위
- **해석**: TPI 1등 = 100%, 꼴찌 ≈ 0%

---

### TPI랭크(전체) / TPI랭크(캠퍼스)
- **정의**: 같은 (시험유형 · 월 · 레벨) 내 TPI 순위 (1 = 최고점)
- 전체 기준 / 캠퍼스 내 기준 두 가지 제공
""")

    st.markdown("---")
    st.markdown("## 🚦 위험등급 분류 기준")

    st.markdown("""
> 모든 위험등급 판단은 **같은 시험유형 · 같은 월 · 같은 레벨** 내에서만 비교합니다.

| 등급 | 조건 | 설명 |
|------|------|------|
| **At-Risk** | 전체 TPI ≤ 하위 20% **AND** B.CV ≤ 하위 20% | 성적도 낮고 과목 간 편차도 큼 → 즉각 개입 필요 |
| **High-Risk** | 전체 TPI ≤ 하위 20% (B.CV 정상) | 전반적 성적 하락 → 학습 지원 필요 |
| **Latent Risk** | B.CV ≤ 하위 20% (TPI 정상) | 과목 편차 과대 → 특정 과목 부진, 잠재 위험 |
| **Local Risk** | 캠퍼스 내 TPI ≤ 하위 20% (타 위험 미해당) | 캠퍼스 내 상대적 부진 |
| **Top Risk** ⭐ | MAG 레벨 전용: 전체 TPI ≥ 상위 20% AND B.CV ≥ 상위 20% | 성적 · 균형 모두 최상위 우수 학생 |
| **Local Top Risk** ⭐ | MAG 레벨 전용: Top Risk 미해당, 캠퍼스 내 TPI · B.CV 모두 ≥ 상위 20% | 캠퍼스 내 최우수 학생 |

> ⭐ Top Risk / Local Top Risk는 **MAG 레벨(MAG1, MAG2 등) 학생에게만 적용**됩니다.

### 판단 순서
1. MAG 레벨이면 → **Top Risk** 확인 → **Local Top Risk** 확인
2. (MAG 포함 전체) **At-Risk** 확인
3. **High-Risk** 확인
4. **Latent Risk** 확인
5. **Local Risk** 확인
6. 해당 없음 → 등급 없음(정상)
""")

    st.markdown("---")
    st.markdown("## 🎨 위험등급 색상 표시")
    _risk_color_guide = {
        "At-Risk":        ("🔴", "#FF4B4B", "즉각 개입 필요"),
        "High-Risk":      ("🟠", "#FF8C00", "성적 하락 주의"),
        "Latent Risk":    ("🟡", "#FFA500", "과목 편차 과대"),
        "Local Risk":     ("🟡", "#FFD700", "캠퍼스 내 부진"),
        "Top Risk":       ("🔵", "#1F4E79", "MAG 최우수 (전체)"),
        "Local Top Risk": ("🔵", "#2E86AB", "MAG 최우수 (캠퍼스)"),
    }
    for grade, (emoji, color, desc) in _risk_color_guide.items():
        st.markdown(
            f'<span style="background:{color};color:white;padding:2px 10px;'
            f'border-radius:4px;font-weight:bold;">{grade}</span> &nbsp; {emoji} {desc}',
            unsafe_allow_html=True,
        )
        st.markdown("")
