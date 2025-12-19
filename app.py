import streamlit as st
import pandas as pd
import plotly.express as px
import json
import os
import time
import streamlit.components.v1 as components 
from st_aggrid import AgGrid, GridOptionsBuilder, ColumnsAutoSizeMode, JsCode 

# 1. 페이지 설정
st.set_page_config(layout="wide", page_title="시장전략팀 통합 분석 v20.0", initial_sidebar_state="expanded")

# KPI 샌드박스 렌더링 함수
def render_kpi_iframe(html_content, is_single=False, is_pair=False):
    pair_class = "pair-mode" if is_pair else ""
    container_style = f"""
    .single-container {{ display: grid; grid-template-columns: repeat(4, 1fr); gap: 35px; background-color: #f0f7ff; border-left: 10px solid #007BFF; border-radius: 15px; padding: 30px; box-shadow: 0 6px 12px rgba(0,0,0,0.08); font-family: 'Pretendard', sans-serif; }}
    .single-title {{ grid-column: 1 / -1; font-size: 26px; font-weight: 800; color: #003a80; margin-bottom: 16px; }}
    .single-item {{ display: flex; flex-direction: column; gap: 8px; border-right: 2px solid #d0e3f7; padding-right: 20px; }}
    .single-item:last-child {{ border-right: none; }}
    .single-label {{ color: #555; font-size: 14px; font-weight: 500; }}
    .single-value {{ font-size: 32px; font-weight: 800; color: #111; line-height: 1.2; }}
    .single-highlight {{ font-size: 34px; font-weight: 900; color: #0066ff; line-height: 1.2; }}
    .single-sub {{ font-size: 12px; color: #888; margin-top: -2px; font-weight: 400; }}
    .compare-container {{ display: flex; gap: 20px; padding: 10px; overflow-x: auto; font-family: 'Pretendard', sans-serif; }}
    .summary-card {{ flex: 1; min-width: 280px; background-color: #ffffff; border-radius: 12px; border-top: 5px solid #007BFF; box-shadow: 0 4px 6px rgba(0,0,0,0.1); padding: 20px; }}
    .card-title {{ font-size: 16px; font-weight: bold; margin-bottom: 12px; border-bottom: 1px solid #eee; padding-bottom: 8px; color: #333; }}
    .card-item {{ display: flex; justify-content: space-between; margin-bottom: 8px; font-size: 14px; align-items: baseline; }}
    .card-highlight {{ color: #007BFF; font-weight: 800; font-size: 18px; }}
    .pair-mode .summary-card {{ min-width: 450px; padding: 30px; }}
    .pair-mode .card-title {{ font-size: 22px; margin-bottom: 18px; }}
    .pair-mode .card-item {{ font-size: 17px; margin-bottom: 12px; }}
    .pair-mode .card-highlight {{ font-size: 24px; }}
    """
    full_html = f"<html><head><style>{container_style}</style></head><body style='background-color: transparent; margin: 0;'><div class='{pair_class}'>{html_content}</div></body></html>"
    components.html(full_html, height=260 if is_single else 320, scrolling=False)

st.markdown("""<h1 style="font-size: 32px; font-weight: 700; margin-bottom: 25px;">📊 시장전략팀 통합 분석 플랫폼 v20.0</h1>""", unsafe_allow_html=True)

USE_COLS = ["유입일자", "센터", "유입경로"]
@st.cache_data
def get_sheets_safe(file_name, file_content, engine): return pd.ExcelFile(file_content, engine=engine).sheet_names

def render_aggrid_final(data, highlight_gn, highlight_ct=None, sort_col="지평"):
    gb = GridOptionsBuilder.from_dataframe(data)
    gb.configure_default_column(cellStyle={'text-align': 'center'}, resizable=True, sortable=True, filter=True)
    gb.configure_grid_options(sortModel=[{"colId": sort_col, "sort": "desc"}])
    gn_json, ct_json = json.dumps(highlight_gn), json.dumps(highlight_ct if highlight_ct else [])
    row_style_jscode = JsCode(f"function(params) {{ if (!params.data) return null; const targetGns = {gn_json}; const targetCts = {ct_json}; if (targetCts.length > 0 && targetCts.includes(params.data.센터)) return {{'backgroundColor': '#FFF9C4', 'color': 'black'}}; if (targetGns.length > 0 && targetGns.includes(params.data.총괄)) return {{'backgroundColor': '#FFF9C4', 'color': 'black'}}; return null; }};")
    gb.configure_grid_options(getRowStyle=row_style_jscode)
    return AgGrid(data, gridOptions=gb.build(), columns_auto_size_mode=ColumnsAutoSizeMode.FIT_CONTENTS, theme='alpine', height=400, allow_unsafe_jscode=True)

# 업로드 상태 세션 관리
if 'data_ready' not in st.session_state: st.session_state.data_ready = False

with st.sidebar:
    if not st.session_state.data_ready:
        st.header("📂 데이터 업로드")
        cust_f = st.file_uploader("고객 로우데이터 (.xlsx / .xlsb)", type=["xlsx", "xlsb"])
        org_f = st.file_uploader("조직 매핑 데이터 (.xlsx / .xlsb)", type=["xlsx", "xlsb"])
        if cust_f and org_f:
            st.session_state.cust_content = cust_f.getvalue()
            st.session_state.org_content = org_f.getvalue()
            st.session_state.cust_name = cust_f.name
            st.session_state.data_ready = True
            st.rerun()
    else:
        if st.button("🔄 데이터 다시 업로드"):
            st.session_state.data_ready = False
            st.rerun()

# 💡 [복구 완료] 사용자 매뉴얼 전문
if not st.session_state.data_ready:
    st.info("👋 분석을 시작하려면 왼쪽 사이드바에서 파일을 업로드해 주세요.")
    st.markdown("""
    ## 📖 시스템 사용 매뉴얼 (User Guide)
    
    ### 1. 분석의 시작 (데이터 업로드)
    - **고객 로우 데이터**: KSS카카오소스관리-고객db관리현황(고객별) 데이터를 업로드합니다 (기준월별 시트 필수) 
    - **조직 매핑 데이터**: 센터별 소속 총괄 및 지국 수가 포함된 데이터를 업로드합니다. (기준월별 시트 필수)
    - 2가지 엑셀 파일은 바이너리 파일을 권장하며 *엑셀 업로드 전 보안해제* 필수입니다.
                
    ### 2. 핵심 개념 이해
    - **📍 조직 편제 기준월**: 데이터의 '분모(지국 수)'를 결정합니다.
        - *예: 2월 성과를 분석할 때, 1월 조직도를 기준으로 보고 싶다면 '1월' 시트를 선택하세요.*
    - **📅 성과 분석 기간**: 실제 추출할 '분자(유입 데이터)'의 범위를 결정합니다.
        - 시작일과 종료일을 지정하면 시트 경계를 넘어 데이터를 통합 분석합니다.
    
    ### 3. 총괄/센터별 비교 분석 활용법 🎯
    - **다중 선택 비교**: 사이드바에서 여러 개의 총괄이나 센터를 선택해 보세요.
        - **총괄별 비교**: 권역별 시장 점유율과 유입 효율의 차이를 한눈에 파악할 수 있습니다.
        - **센터별 비교**: 동일 총괄 내 센터 간 성과 편차를 확인하여 우수 사례를 발굴할 수 있습니다.
    - **🔥 Pair-mode (가독성 모드)**: 
        - 단 2개의 조직만 선택할 경우, 비교 효율을 극대화하기 위해 **카드와 글자 크기가 자동으로 커집니다.**
        - 1:1 라이벌 매치나 전략적 집중 비교에 효과적입니다.
    
    ### 4. 화면 보는 법 및 기능 가이드
    - **KPI 카드**: 지국 수 대비 유입 효율(지평)을 한눈에 보여줍니다.
    - **상세 랭킹**: 테이블 상단 컬럼을 클릭하여 유입량이나 지평 순위를 즉시 정렬할 수 있습니다.
    """)
else:
    try:
        engine_c = 'pyxlsb' if st.session_state.cust_name.endswith('.xlsb') else 'calamine'
        cache_fn = f"cache_{st.session_state.cust_name}.parquet"
        
        if os.path.exists(cache_fn):
            df_all = pd.read_parquet(cache_fn)
            sheet_names = get_sheets_safe(st.session_state.cust_name, st.session_state.cust_content, engine_c)
        else:
            sheet_names = get_sheets_safe(st.session_state.cust_name, st.session_state.cust_content, engine_c)
            with st.status("🚀 분석 엔진 빌드 중...", expanded=True) as status:
                combined = [pd.read_excel(st.session_state.cust_content, sheet_name=s, engine=engine_c, usecols=USE_COLS).assign(원본시트=str(s).strip()) for s in sheet_names]
                df_all = pd.concat(combined, ignore_index=True)
                df_all.to_parquet(cache_fn, index=False)
                status.update(label="✅ 엔진 가동 완료", state="complete", expanded=False)

        df_all["timestamp"] = pd.to_datetime(df_all["유입일자"].astype(str).str.replace("시", "").str.replace(".", "-"), errors='coerce')
        df_all = df_all.dropna(subset=["timestamp"]).copy()
        df_all["날짜"] = df_all["timestamp"].dt.date
        df_all["hour"] = df_all["timestamp"].dt.hour
        df_all["요일"] = df_all["timestamp"].dt.day_name().map({'Monday':'월','Tuesday':'화','Wednesday':'수','Thursday':'목','Friday':'금','Saturday':'토','Sunday':'일'})

        with st.sidebar:
            st.divider()
            st.header("⚙️ 분석 설정")
            base_month = st.selectbox("📍 조직 편제 기준월", sheet_names, index=len(sheet_names)-1)
            d_range = st.date_input("📅 분석 조회 기간", [df_all['날짜'].min(), df_all['날짜'].max()])
            df_o_current = pd.read_excel(st.session_state.org_content, sheet_name=base_month)
            df_o_current.columns = [str(c).strip() for c in df_o_current.columns]
            df_o_current["지국"] = pd.to_numeric(df_o_current["지국"], errors='coerce').fillna(0)
            selected_gn = st.multiselect("🎯 총괄 선택", sorted(df_o_current["총괄"].unique().tolist()))
            temp_org = df_o_current[df_o_current["총괄"].isin(selected_gn)] if selected_gn else df_o_current
            selected_ct = st.multiselect("🏬 센터 선택", sorted(temp_org["센터"].unique().tolist()))
            target_org_info = temp_org[temp_org["센터"].isin(selected_ct)] if selected_ct else temp_org

        df_filtered = df_all[(df_all["날짜"] >= d_range[0]) & (df_all["날짜"] <= d_range[1])]
        df_merged = pd.merge(df_filtered, df_o_current[['센터', '총괄', '지국']], on='센터', how='inner')
        df_target = df_merged[df_merged['센터'].isin(target_org_info['센터'])]

        if not df_target.empty:
            st.subheader(f"🏁 분석 성과 요약 ({d_range[0]} ~ {d_range[1]})")
            st.caption(f"조직 편제 기준: {base_month} 시트 적용")
            
            display_list = (selected_ct if selected_ct else selected_gn) if (selected_ct or selected_gn) else ["전체"]
            is_single, is_pair = (len(display_list) == 1), (len(display_list) == 2)
            group_col = "센터" if selected_ct else ("총괄" if selected_gn else None)
            
            card_html = '<div class="compare-container">' if not is_single else ""
            for org in display_list:
                t_df = df_target if org == "전체" else df_target[df_target[group_col]==org]
                t_meta = target_org_info if org == "전체" else target_org_info[target_org_info[group_col]==org]
                jg_sum, total_src = int(t_meta['지국'].sum()), len(t_df)
                eff = total_src / jg_sum if jg_sum > 0 else 0
                peak_day, peak_hour = t_df['요일'].mode().iloc[0] if not t_df['요일'].mode().empty else "-", t_df['hour'].mode().iloc[0] if not t_df['hour'].mode().empty else "-"
                
                sub_text = f"<span class='single-sub'>※ {base_month} 기준</span>"
                if is_single:
                    card_html = f"""<div class="single-container"><div class="single-title">🏢 {org} 성과 요약</div><div class="single-item"><span class="single-label">총 지국수</span><span class="single-value">{jg_sum:,}</span>{sub_text}</div><div class="single-item"><span class="single-label">총 유입</span><span class="single-value">{total_src:,}건</span></div><div class="single-item"><span class="single-label">지평(효율)</span><span class="single-highlight">{eff:.2f}</span></div><div class="single-item"><span class="single-label">피크 타임</span><span class="single-value">{peak_day}요일 · {peak_hour}시</span></div></div>"""
                else:
                    card_html += f"""<div class="summary-card"><div class="card-title">[ {org} ]</div><div class="card-item"><span class="card-label">총 유입</span><span class="card-value">{total_src:,}건</span></div><div class="card-item"><span class="card-label">지평</span><span class="card-highlight">{eff:.2f}</span></div><div class="card-item"><span class="card-label">피크</span><span class="card-value">{peak_day}요일·{peak_hour}시</span></div></div>"""
            if not is_single: card_html += '</div>'
            render_kpi_iframe(card_html, is_single, is_pair)

            st.divider()
            v_t = st.tabs(["📉 일자별 추이 분석", "📢 채널 유입 분석", "🔥 요일/시간별 히트맵", "🏆 조직별 소스 유입 랭킹"])
            
            with v_t[0]: 
                st.info("**🧐 어떻게 보나요?**: 일자별 소스 유입량의 변화를 선 그래프로 확인합니다.  \n**💡 인사이트**: 유입이 튀는 날짜는 특정 프로모션이나 외부 이슈의 결과일 확률이 높습니다.")
                st.plotly_chart(px.line(df_target.groupby(["날짜", group_col]).size().reset_index(name="건") if group_col else df_target.groupby("날짜").size().reset_index(name="건"), x="날짜", y="건", color=group_col, markers=True), use_container_width=True)
            
            with v_t[1]: 
                st.info("**🧐 어떻게 보나요?**: 어떤 유입 경로(소스)가 가장 효과적이었는지 비교합니다.  \n**💡 인사이트**: 설정한 기간 내 유입 비중이 높은 채널을 확인할 수 있습니다(박람회 등).  \n**⚙️ 기능 안내**: 마우스를 막대 위에 올리면(Hover) 정확한 건수를 수치로 확인할 수 있습니다.")
                st.plotly_chart(px.bar(df_target.groupby(["유입경로", group_col]).size().reset_index(name="건") if group_col else df_target.groupby("유입경로").size().reset_index(name="건"), x="건", y="유입경로", color=group_col, barmode="group", orientation='h', text_auto=True), use_container_width=True)
            
            with v_t[2]: 
                st.info("**🧐 어떻게 보나요?**: 요일과 시간의 교차점을 통해 소스가 유입되는 '골든 타임'을 확인 가능합니다.  \n**💡 인사이트**: 색상이 짙은 영역에 개척 영업 활동이 이루어지고 있다고 볼 수 있으며, 이때 맞춰 프로모션 수립이 가능합니다.  \n**⚙️ 기능 안내**: 세로축은 요일, 가로축은 24시간을 나타내며 실시간 패턴 분석을 지원합니다.")
                st.plotly_chart(px.density_heatmap(df_target.groupby(["요일", "hour"]).size().reset_index(name="건"), x="hour", y="요일", z="건", text_auto=True, color_continuous_scale="Blues", category_orders={"요일": ['월','화','수','목','금','토','일']}), use_container_width=True)
            
            with v_t[3]:
                st.info("**🧐 어떻게 보나요?**: 총괄 및 센터별 세부 성과를 테이블 형태로 분석합니다.  \n**💡 인사이트**: 유입 절대량보다 조직 활동 인원에 편차가 있어 '지평' 순위가 높은 조직 위주로 살펴봅니다.  \n**⚙️ 기능 안내**: 각 컬럼을 클릭하면 오름차순/내림차순 정렬이 가능하며, 왼쪽 조직을 선택했을 경우 해당되는 조직은 노란색으로 강조됩니다.")
                sub_rank = st.tabs(["🏢 총괄 순위", "🏬 센터별 순위"])
                with sub_rank[0]: render_aggrid_final(pd.merge(df_merged.groupby('총괄').size().reset_index(name='유입'), df_o_current.groupby('총괄')['지국'].sum().reset_index(name='지국수'), on='총괄').assign(지평=lambda x: (x['유입']/x['지국수']).round(2)), selected_gn)
                with sub_rank[1]: render_aggrid_final(df_merged.groupby(['총괄', '센터']).agg(유입=('유입경로','count'), 지국수=('지국','max')).reset_index().assign(지평=lambda x: (x['유입']/x['지국수']).round(2)), selected_gn, selected_ct)

    except Exception as e:
        st.error(f"🚨 시스템 오류 확인: {e}")
