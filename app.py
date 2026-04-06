"""
오성AIT 홈택스 전자세금계산서 일괄발행 시스템
ERP 매출 파일 업로드 → 마스터 자동 매핑 → 미리보기 → 엑셀 다운로드
"""

import streamlit as st
import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
import os, io, re, urllib.request
from datetime import datetime

st.set_page_config(
    page_title="오성AIT 세금계산서",
    page_icon="🧾",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""<style>
.block-container{padding-top:1.2rem !important}
.stApp{background:#f8f9fb;color:#1e293b}
section[data-testid="stSidebar"]{background:#fff;border-right:1px solid #e2e8f0}
#MainMenu,footer,header{display:none !important}
.stButton>button{border-radius:8px !important;font-weight:600 !important}
.stDownloadButton>button{border-radius:8px !important;font-weight:700 !important;
  font-size:15px !important;padding:12px !important}
.stTabs [data-baseweb="tab-list"]{background:#f1f5f9;border-radius:10px;padding:4px;gap:2px}
.stTabs [data-baseweb="tab"]{border-radius:8px;color:#64748b;font-size:14px;
  font-weight:600;padding:10px 20px}
.stTabs [aria-selected="true"]{background:#fff !important;color:#1e293b !important;
  box-shadow:0 1px 4px rgba(0,0,0,.08) !important}
[data-testid="stMetric"]{background:#fff !important;border:1px solid #e2e8f0 !important;
  border-radius:10px !important;padding:14px !important}
</style>""", unsafe_allow_html=True)

SUPPLIER = {
    "biz_no": "4031409207", "name": "오성AIT", "ceo": "이후종",
    "addr":   "전북 익산시 오산면 오산로 151",
    "type":   "도/소매", "item": "철물/건재", "email": "ohsungait@naver.com",
}
FIXED_ITEM = "철물 외"
FIXED_NOTE = "농협(오성AIT) 351-0830-5542-93"
GITHUB_RAW = (
    "https://raw.githubusercontent.com/ohsungait-droid/ohsung-auto/main/"
    "%EC%98%A4%EC%84%B1%EC%A0%84%EC%9E%90%EC%84%B8%EA%B8%88%EC%97%85%EB%A1%9C%EB%93%9C.xlsm"
)

for k, v in {
    "master_index": None, "master_loaded": False,
    "preview_df": None, "write_date": datetime.now().strftime("%Y%m%d"),
}.items():
    if k not in st.session_state:
        st.session_state[k] = v

def norm_regno(val) -> str:
    if val is None: return ""
    return re.sub(r"[-\s]", "", str(val)).strip()

def norm_name(val) -> str:
    if val is None: return ""
    return re.sub(r"\s+", "", str(val)).strip().lower()

def parse_number(val) -> int:
    if val is None or val == "": return 0
    try: return int(float(str(val).replace(",", "")))
    except: return 0

def load_master(buf: bytes) -> dict:
    wb = load_workbook(io.BytesIO(buf), read_only=True, keep_vba=True)
    ws = wb["마스터시트"] if "마스터시트" in wb.sheetnames else wb.active
    by_regno, by_name = {}, {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or not row[0]: continue
        rn = norm_regno(row[0])
        if not rn: continue
        rec = {
            "biz_no": rn, "name": str(row[1] or "").strip(),
            "ceo":    str(row[2] or "").strip(), "addr": str(row[3] or "").strip(),
            "type":   str(row[4] or "").strip(), "item": str(row[5] or "").strip(),
            "email1": str(row[6] or "").strip(), "email2": str(row[7] or "").strip(),
        }
        by_regno[rn] = rec
        nn = norm_name(rec["name"])
        if nn: by_name[nn] = rec
    wb.close()
    return {"byRegNo": by_regno, "byName": by_name}

def lookup_master(index: dict, regno: str, name: str):
    if regno:
        found = index["byRegNo"].get(norm_regno(regno))
        if found: return found
    if name:
        found = index["byName"].get(norm_name(name))
        if found: return found
    return None

def parse_erp(buf: bytes) -> pd.DataFrame:
    df_raw = pd.read_excel(io.BytesIO(buf), header=None)
    header_idx = 0
    for i in range(min(5, len(df_raw))):
        if df_raw.iloc[i].notna().sum() > df_raw.iloc[header_idx].notna().sum():
            header_idx = i
    headers = [str(h or "").strip().lower() for h in df_raw.iloc[header_idx]]
    df = df_raw.iloc[header_idx+1:].copy()
    df.columns = headers
    df = df.reset_index(drop=True)

    def fc(*kws):
        for kw in kws:
            for h in headers:
                if kw in h: return h
        return None

    col_client = fc("거래처명","거래처","상호","업체명")
    col_regno  = fc("사업자번호","사업자등록","등록번호")
    col_supply = fc("공급가액","공급가","공급액")
    col_tax    = fc("부가세","세액","vat")
    col_date   = fc("거래일","일자","날짜","date","매출일")
    col_unit   = fc("단가")

    rows = []
    for _, row in df.iterrows():
        if col_unit and parse_number(row.get(col_unit, 0)) == 0: continue
        supply = parse_number(row.get(col_supply, 0)) if col_supply else 0
        tax    = parse_number(row.get(col_tax,    0)) if col_tax    else 0
        if supply == 0 and tax == 0: continue
        rows.append({
            "거래처명":   str(row.get(col_client,"") or "").strip() if col_client else "",
            "사업자번호": norm_regno(row.get(col_regno,"")) if col_regno else "",
            "공급가액":   supply, "부가세": tax,
            "날짜":       str(row.get(col_date,"") or "").strip() if col_date else "",
        })
    if not rows: return pd.DataFrame()
    df_out = pd.DataFrame(rows)
    key = df_out["사업자번호"].where(df_out["사업자번호"] != "", df_out["거래처명"])
    df_out["_key"] = key
    agg = df_out.groupby("_key", sort=False).agg(
        거래처명=("거래처명","first"), 사업자번호=("사업자번호","first"),
        공급가액=("공급가액","sum"), 부가세=("부가세","sum"), 날짜=("날짜","last"),
    ).reset_index(drop=True)
    return agg

def build_preview(erp_df: pd.DataFrame, master_index: dict, write_date: str) -> pd.DataFrame:
    dd = write_date[6:8] if len(write_date) >= 8 else write_date
    rows = []
    for _, row in erp_df.iterrows():
        m = lookup_master(master_index, row["사업자번호"], row["거래처명"])
        rows.append({
            "매칭": "✅" if m else "⚠️",
            "ERP 거래처명": row["거래처명"],
            "마스터 거래처명": m["name"] if m else "—",
            "사업자번호": m["biz_no"] if m else row["사업자번호"] or "—",
            "이메일": m["email1"] if m else "—",
            "공급가액": int(row["공급가액"]),
            "부가세":   int(row["부가세"]),
            "_matched": bool(m), "_master": m,
            "_erp_client": row["거래처명"],
            "_write_date": write_date, "_item_day": dd,
        })
    return pd.DataFrame(rows)

def make_hometax_xlsx(preview_df: pd.DataFrame, include_unmatched: bool = False) -> bytes:
    filtered = preview_df if include_unmatched else preview_df[preview_df["_matched"]]
    if filtered.empty: return b""
    wb = Workbook()
    ws = wb.active
    ws.title = "엑셀업로드양식"
    thin = Side(style="thin", color="BFBFBF")
    bdr  = Border(left=thin,right=thin,top=thin,bottom=thin)
    headers = [
        "전자(세금)계산서 종류\n(01:일반, 02:영세율)","작성일자",
        "공급자 등록번호\n(\"-\" 없이 입력)","공급자\n 종사업장번호",
        "공급자 상호","공급자 성명","공급자 사업장주소","공급자 업태","공급자 종목","공급자 이메일",
        "공급받는자 등록번호\n(\"-\" 없이 입력)","공급받는자 \n종사업장번호",
        "공급받는자 상호 ","공급받는자 성명","공급받는자 사업장주소",
        "공급받는자 업태","공급받는자 종목","공급받는자 이메일1","공급받는자 이메일2",
        "공급가액\n합계","세액\n합계","비고",
        "일자1\n(2자리, 작성년월 제외)","품목1","규격1","수량1","단가1","공급가액1","세액1","품목비고1",
        "일자2\n(2자리, 작성년월 제외)","품목2","규격2","수량2","단가2","공급가액2","세액2","품목비고2",
        "일자3\n(2자리, 작성년월 제외)","품목3","규격3","수량3","단가3","공급가액3","세액3","품목비고3",
        "일자4\n(2자리, 작성년월 제외)","품목4","규격4","수량4","단가4","공급가액4","세액4","품목비고4",
        "현금","수표","어음","외상미수금","영수(01),\n청구(02)",
    ]
    ws.row_dimensions[6].height = 36
    for ci, h in enumerate(headers, 1):
        cell = ws.cell(row=6, column=ci, value=h)
        cell.font = Font(name="맑은 고딕", size=9, bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = bdr
        if 3 <= ci <= 10:
            cell.fill = PatternFill("solid", fgColor="D9E1F2")
        elif 11 <= ci <= 19:
            cell.fill = PatternFill("solid", fgColor="FCE4D6")
        else:
            cell.fill = PatternFill("solid", fgColor="F2F2F2")

    for ri, (_, row) in enumerate(filtered.iterrows(), 7):
        m = row["_master"] or {}
        vals = [
            "01", row["_write_date"], SUPPLIER["biz_no"], "",
            SUPPLIER["name"], SUPPLIER["ceo"], SUPPLIER["addr"],
            SUPPLIER["type"], SUPPLIER["item"], SUPPLIER["email"],
            m.get("biz_no","") if row["_matched"] else row.get("사업자번호",""), "",
            m.get("name", row["_erp_client"]),
            m.get("ceo",""), m.get("addr",""), m.get("type",""), m.get("item",""),
            m.get("email1",""), m.get("email2",""),
            row["공급가액"], row["부가세"], FIXED_NOTE,
            row["_item_day"], FIXED_ITEM, "","","",
            row["공급가액"], row["부가세"], "",
            "","","","","","","","",
            "","","","","","","","",
            "","","","","","","","",
            "","","","","02",
        ]
        ws.row_dimensions[ri].height = 18
        h_bdr = Border(
            left=Side(style="hair",color="BFBFBF"), right=Side(style="hair",color="BFBFBF"),
            top=Side(style="hair",color="BFBFBF"),  bottom=Side(style="hair",color="BFBFBF"),
        )
        for ci, v in enumerate(vals, 1):
            cell = ws.cell(row=ri, column=ci, value=v)
            cell.font = Font(name="맑은 고딕", size=9)
            cell.alignment = Alignment(horizontal="left", vertical="center")
            cell.border = h_bdr

    ws.freeze_panes = "A7"
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()

# ═══════════════ 사이드바 ═══════════════
with st.sidebar:
    logo_path = "ICON (1).png"
    if os.path.exists(logo_path):
        c1, c2 = st.columns([1,2])
        with c1: st.image(logo_path, width=52)
        with c2:
            st.markdown("<div style='padding-top:8px'>"
                "<b style='font-size:14px;color:#1e293b'>오성AIT</b><br>"
                "<span style='font-size:10px;color:#94a3b8'>홈택스 세금계산서</span></div>",
                unsafe_allow_html=True)
    else:
        st.markdown("### 🧾 오성AIT")

    st.markdown("---")

    if not st.session_state.master_loaded:
        with st.spinner("☁️ 마스터시트 로드 중..."):
            try:
                token = ""
                try: token = st.secrets.get("GITHUB_TOKEN","")
                except: pass
                req = urllib.request.Request(
                    GITHUB_RAW,
                    headers={"Authorization": f"token {token}"} if token else {}
                )
                with urllib.request.urlopen(req) as resp:
                    buf = resp.read()
                idx = load_master(buf)
                st.session_state.master_index  = idx
                st.session_state.master_loaded = True
            except Exception as e:
                st.warning(f"⚠️ 자동 로드 실패")

    if st.session_state.master_loaded:
        cnt = len(st.session_state.master_index["byRegNo"])
        st.success(f"✅ 마스터 {cnt}개 거래처")
        if st.button("🔄 새로고침", use_container_width=True):
            st.session_state.master_loaded = False
            st.rerun()
    else:
        st.markdown("**홈택스 파일 수동 업로드**")
        ht_up = st.file_uploader("xlsm", type=["xlsm","xlsx"],
                                  key="ht_manual", label_visibility="collapsed")
        if ht_up:
            idx = load_master(ht_up.read())
            st.session_state.master_index  = idx
            st.session_state.master_loaded = True
            st.rerun()

    st.markdown("---")
    st.markdown("**📅 작성일자**")
    write_date = st.text_input("작성일자", value=st.session_state.write_date,
                               placeholder="YYYYMMDD", label_visibility="collapsed")
    st.session_state.write_date = write_date

    st.markdown("---")
    st.caption("☁️ Cloud 버전 · 데이터 저장 없음")

# ═══════════════ 메인 ═══════════════
st.markdown(
    "<h1 style='margin:0;font-size:24px;font-weight:900;color:#1e293b'>"
    "🧾 홈택스 전자세금계산서 일괄발행</h1>"
    "<p style='margin:4px 0 16px;font-size:13px;color:#64748b'>"
    "ERP 파일 업로드 → 마스터 자동 매핑 → 미리보기 → 엑셀 다운로드</p>",
    unsafe_allow_html=True
)

st.markdown("#### 📊 STEP 1 — ERP 매출 파일 업로드")
erp_upload = st.file_uploader(
    "이번 달 ERP 매출 파일 (.xlsx)",
    type=["xlsx","xls"],
    help="거래처명, 공급가액, 부가세가 포함된 매출 파일"
)

if erp_upload and st.session_state.master_loaded:
    with st.spinner("파일 분석 중..."):
        erp_df = parse_erp(erp_upload.read())

    if erp_df.empty:
        st.error("❌ 데이터를 찾을 수 없어요.")
        st.stop()

    preview_df = build_preview(erp_df, st.session_state.master_index,
                               st.session_state.write_date)
    st.session_state.preview_df = preview_df

    st.markdown("---")
    st.markdown("#### 📋 STEP 2 — 매핑 결과 미리보기")

    matched   = int(preview_df["_matched"].sum())
    unmatched = int((~preview_df["_matched"]).sum())
    total_sup = int(preview_df["공급가액"].sum())

    c1,c2,c3,c4 = st.columns(4)
    c1.metric("전체 거래처",  f"{len(preview_df)}개")
    c2.metric("✅ 매핑 완료", f"{matched}개")
    c3.metric("⚠️ 미매핑",   f"{unmatched}개")
    c4.metric("공급가액 합계", f"{total_sup:,.0f}원")

    if unmatched > 0:
        unmatch_list = preview_df[~preview_df["_matched"]]["ERP 거래처명"].tolist()
        st.warning(f"⚠️ 마스터 미등록 **{unmatched}개** — 홈택스 파일에서 제외됩니다.")
        with st.expander("미등록 거래처 목록"):
            for nm in unmatch_list:
                st.markdown(f"- {nm}")
    else:
        st.success("✅ 모든 거래처 매핑 완료!")

    disp = preview_df[[
        "매칭","ERP 거래처명","마스터 거래처명","사업자번호","이메일","공급가액","부가세"
    ]].copy()
    disp["공급가액"] = disp["공급가액"].apply(lambda x: f"{x:,}원")
    disp["부가세"]   = disp["부가세"].apply(lambda x: f"{x:,}원")
    st.dataframe(disp, use_container_width=True, hide_index=True, height=420)

    st.markdown("---")
    st.markdown("#### 📥 STEP 3 — 홈택스 엑셀 다운로드")

    include_unmatched = st.checkbox("미매핑 거래처도 포함 (사업자번호 없이)", value=False)
    target = len(preview_df) if include_unmatched else matched

    if target == 0:
        st.warning("다운로드할 데이터가 없습니다.")
    else:
        xlsx_bytes = make_hometax_xlsx(preview_df, include_unmatched)
        if xlsx_bytes:
            fname = f"홈택스_세금계산서_{st.session_state.write_date[:6]}.xlsx"
            st.download_button(
                label=f"📥 홈택스 엑셀 다운로드  ({target}개 업체)",
                data=xlsx_bytes, file_name=fname,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True, type="primary",
            )
            st.caption("💡 홈택스 → 전자세금계산서 → 일괄발행 → 파일업로드 에서 이 파일을 올려주세요.")

elif erp_upload and not st.session_state.master_loaded:
    st.warning("⚠️ 사이드바에서 홈택스 파일을 먼저 업로드해주세요.")
else:
    st.markdown("""
    <div style='background:#fff;border:2px dashed #e2e8f0;border-radius:14px;
    padding:48px;text-align:center;color:#94a3b8;margin-top:20px'>
    <div style='font-size:52px'>📊</div>
    <div style='font-size:17px;font-weight:700;color:#64748b;margin:14px 0 8px'>
    ERP 매출 파일을 업로드하세요</div>
    <div style='font-size:13px'>거래처명 · 공급가액 · 부가세가 포함된 엑셀 파일</div>
    </div>
    """, unsafe_allow_html=True)
