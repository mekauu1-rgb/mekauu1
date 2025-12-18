import streamlit as st
import pandas as pd
import plotly.express as px

# 1. 페이지 설정 및 디자인 (글자 크기 강화)
st.set_page_config(layout="wide", page_title="시장전략팀 유입 분석 리포트")

st.markdown("""
    <style>
    html, body, [class*="st-"] { font-size: 1.2rem; }
    h1 { font-size: 3.0rem !important; color: #1E3A8A; font-weight: 800; }
    h2 { font-size: 2.4rem !important; border-bottom: 3px solid #2563EB; padding-bottom: 10px; margin-top: 40px; }
    h3 { font-size: 2.0rem !important; color: #1F2937; margin-top: 30px; }
    .stMetric label { font-size: 1.4rem !important; font-weight: bold; }
    .stMetric div { font-size: 2.6rem !important; color: #2563EB; font-weight: 700; }
    .insight-box { background-color: #F8FAFC; padding: 25px; border-radius: 15px; border: 1px solid #E2E8F0; border-left: 8px solid #2563EB; }
    .winner-text { color: #E11D48; font-weight: 800; font-size: 1.6rem; }
    </style>
    """, unsafe_allow_html=True)

st.title("📊 고객 유입 통합 심층 비교 리포트")

# 2. 파일 업로드
col_up1, col_up2 = st.columns(2)
with col_up1:
    customer_file = st.file_uploader("📂 고객 로우데이터 업로드 (xlsx)", type=["xlsx"])
with col_up2:
    org_file = st.file_uploader("📂 조직 매핑 데이터 업로드 (xlsx)", type=["xlsx"])

if not customer_file or not org_file:
    st.info("💡 분석을 위해 두 개의 엑셀 파일을 모두 업로드해 주세요.")
    st.stop()

# 3. 데이터 로드 및 전처리
@st.cache_data
def load_and_preprocess(file1, file2):
    try:
        df_c = pd.read_excel(file1, engine='openpyxl')
        df_o = pd.read_excel(file2, engine='openpyxl')
        df_c.columns = df_c.columns.str.strip()
        df_o.columns = df_o.columns.str.strip()

        # 날짜 파싱
        def parse_dt(x):
            if pd.isna(x): return pd.NaT
            s = str(x).replace("시", "").strip()
            for fmt in ("%Y.%m.%d %H", "%Y-%m-%d %H", "%Y.%m.%d", "%Y-%m-%d"):
                try: return pd.to_datetime(s, format=fmt)
                except: continue
            return pd.to_datetime(s, errors='coerce')

        df_c["timestamp"] = df_c["유입일자"].apply(parse_dt)
        df_c = df_c.dropna(subset=["timestamp"])
        df_c["date"] = df_c["timestamp"].dt.date
        df_c["hour"] = df_c["timestamp"].dt.hour
        weekday_map = {0: "월", 1: "화", 2: "수", 3: "목", 4: "금", 5: "토", 6: "일"}
        df_c["요일"] = df_c["timestamp"].dt.weekday.map(weekday_map)
        
        # 조직 정보 결합 (총괄, 지국 포함)
        if "센터" in df_c.columns and "센터" in df_o.columns:
            # 센터별 지국 합산 데이터 생성
            df_o_sub = df_o.groupby("센터").agg({"총괄": "first", "지국": "sum"}).reset_index()
            df_c = df_c.merge(df_o_sub, on="센터", how="left")
        
        return df_c, df_o_sub
    except Exception as e:
        st.error(f"데이터 처리 오류: {e}")
        return None, None

df, org_summary = load_and_preprocess(customer_file, org_file)

if df is not None:
    # 4. 사이드바 설정
    st.sidebar.header("🔍 통합 조회 설정")
    min_date, max_date = df["date"].min(), df["date"].max()
    date_range = st.sidebar.date_input("📅 조회 기간", value=(min_date, max_date))
    customer_type = st.sidebar.radio("👤 고객 구분", ["전체", "신규"], horizontal=True)
    scope_option = st.sidebar.selectbox("🏢 분석 단위 선택", ["전사", "총괄", "센터"])
    
    selected_targets = []
    if scope_option != "전사":
        targets = sorted(df[scope_option].dropna().unique().tolist())
        selected_targets = st.sidebar.multiselect(f"🔎 비교 대상({scope_option})", targets, default=targets[:2] if len(targets)>1 else targets)

    selected_channel = st.sidebar.selectbox("🔗 유입채널", ["전체"] + sorted(df["유입경로"].dropna().unique().tolist()))
    period_filter = st.sidebar.radio("📆 요일 구분", ["전체", "평일", "주말"], horizontal=True)

    # 데이터 필터링 적용
    fdf = df.copy()
    if len(date_range) == 2: fdf = fdf[(fdf["date"] >= date_range[0]) & (fdf["date"] <= date_range[1])]
    if customer_type == "신규": fdf = fdf[fdf["신규"] == "Y"]
    if scope_option != "전사" and selected_targets: fdf = fdf[fdf[scope_option].isin(selected_targets)]
    if selected_channel != "전체": fdf = fdf[fdf["유입경로"] == selected_channel]
    if period_filter == "평일": fdf = fdf[fdf["timestamp"].dt.weekday < 5]
    elif period_filter == "주말": fdf = fdf[fdf["timestamp"].dt.weekday >= 5]

    # 5. 메인 리포트 요약
    st.header("✨ 주요 유입 요약")
    if scope_option != "전사" and selected_targets:
        m_cols = st.columns(len(selected_targets))
        for i, target in enumerate(selected_targets):
            t_df = fdf[fdf[scope_option] == target]
            # 지국 계산
            target_jg = org_summary[org_summary[scope_option == "총괄" and "총괄" or "센터"] == target]["지국"].sum() if scope_option in ["총괄", "센터"] else 0
            with m_cols[i]:
                st.subheader(f"🏢 {target}")
                st.metric("총 유입", f"{len(t_df):,}건")
                st.write(f"👥 **지국 수:** {target_jg:,}개")
                if not t_df.empty:
                    st.write(f"🔝 **최다 유입:** {t_df['요일'].mode()[0]}요일 / {t_df['hour'].mode()[0]}시")
    else:
        m1, m2, m3 = st.columns(3)
        m1.metric("전체 총 유입 건수", f"{len(fdf):,}건")
        m2.metric("전체 최다 요일", f"{fdf['요일'].mode()[0]}요일")
        m3.metric("전체 최다 시간대", f"{fdf['hour'].mode()[0]}시")

    st.markdown("---")

    # 6. 차트 분석 섹션
    col1, col2 = st.columns(2)
    day_order = ["월", "화", "수", "목", "금", "토", "일"]
    
    with col1:
        st.subheader("📅 요일별 유입 패턴")
        group_cols = [scope_option, "요일"] if scope_option != "전사" else ["요일"]
        day_data = fdf.groupby(group_cols).size().reset_index(name="건수")
        fig_day = px.bar(day_data, x="요일", y="건수", color=scope_option if scope_option != "전사" else None, 
                         barmode="group", category_orders={"요일": day_order}, text_auto=True)
        st.plotly_chart(fig_day, use_container_width=True)

    with col2:
        st.subheader("🕒 시간대별 유입 흐름")
        group_cols_h = [scope_option, "hour"] if scope_option != "전사" else ["hour"]
        hour_data = fdf.groupby(group_cols_h).size().reset_index(name="건수")
        if scope_option != "전사":
            fig_hour = px.line(hour_data, x="hour", y="건수", color=scope_option, markers=True)
        else:
            fig_hour = px.bar(hour_data, x="hour", y="건수", text_auto=True)
        fig_hour.update_layout(xaxis=dict(tickmode='linear', dtick=1))
        st.plotly_chart(fig_hour, use_container_width=True)

    st.markdown("---")

    # 7. 상세 분석 (채널 비중 & 히트맵) - 전사 모드 포함
    st.subheader("🔗 채널 비중 및 시간대 집중도")
    if scope_option != "전사" and len(selected_targets) >= 1:
        for target in selected_targets:
            t_df = fdf[fdf[scope_option] == target]
            c_left, c_right = st.columns([4, 6])
            with c_left:
                fig_pie = px.pie(t_df["유입경로"].value_counts().reset_index(), values="count", names="유입경로", hole=0.3, title=f"[{target}] 채널 비중")
                st.plotly_chart(fig_pie, use_container_width=True)
            with c_right:
                heat = t_df.pivot_table(index="요일", columns="hour", values="timestamp", aggfunc="count", fill_value=0).reindex(day_order)
                fig_heat = px.imshow(heat, text_auto=True, aspect="auto", color_continuous_scale="YlOrRd", title=f"[{target}] 시간대 집중도")
                st.plotly_chart(fig_heat, use_container_width=True)
    else:
        # 전사 모드일 때 전체 채널 비중 및 히트맵 출력
        c_left, c_right = st.columns([4, 6])
        with c_left:
            fig_pie = px.pie(fdf["유입경로"].value_counts().reset_index(), values="count", names="유입경로", hole=0.3, title="[전사] 채널 비중")
            st.plotly_chart(fig_pie, use_container_width=True)
        with c_right:
            heat = fdf.pivot_table(index="요일", columns="hour", values="timestamp", aggfunc="count", fill_value=0).reindex(day_order)
            fig_heat = px.imshow(heat, text_auto=True, aspect="auto", color_continuous_scale="YlOrRd", title="[전사] 시간대 집중도")
            st.plotly_chart(fig_heat, use_container_width=True)

    # ---------------------------------------------------------
    # 8. 🏆 시장전략팀의 심층 비교 분석 결론
    # ---------------------------------------------------------
    if scope_option != "전사" and len(selected_targets) >= 2:
        st.markdown("---")
        st.header("🧐 시장전략팀의 심층 비교 분석 결론")
        
        # 데이터 집계
        summary = fdf.groupby(scope_option).size().reset_index(name="total").sort_values(by="total", ascending=False)
        winner = summary.iloc[0]
        runner_up = summary.iloc[1]
        
        # 지국 데이터 매핑 로직 (총괄/센터 구분)
        jg_col = "총괄" if scope_option == "총괄" else "센터"
        winner_jg = org_summary[org_summary[jg_col] == winner[scope_option]]["지국"].sum()
        runner_jg = org_summary[org_summary[jg_col] == runner_up[scope_option]]["지국"].sum()

        # 인당(지국당) 유입 효율
        winner_eff = winner['total'] / winner_jg if winner_jg > 0 else 0
        runner_eff = runner_up['total'] / runner_jg if runner_jg > 0 else 0

        # 유입 취약 시간대/채널 분석
        winner_df = fdf[fdf[scope_option] == winner[scope_option]]
        runner_df = fdf[fdf[scope_option] == runner_up[scope_option]]
        
        # 우세 조직 안내
        st.markdown(f"### 🏆 전반적인 소스 볼륨 면에서 <span class='winner-text'>{winner[scope_option]}</span>이(가) 우세한 지표입니다.", unsafe_allow_html=True)
        
        with st.container():
            st.markdown(f"""
            <div class="insight-box">
            <h4>🔍 심층 분석 결과 보고서</h4>
            <ul style="line-height: 1.8;">
                <li><b>1. 조직 규모 및 유입량 대조:</b> <br>
                    - <b>{winner[scope_option]}</b>: 지국 수 <b>{winner_jg:,}개</b> / 총 유입 <b>{winner['total']:,}건</b> (지국당 <b>{winner_eff:.1f}건</b>) <br>
                    - <b>{runner_up[scope_option]}</b>: 지국 수 <b>{runner_jg:,}개</b> / 총 유입 <b>{runner_up['total']:,}건</b> (지국당 <b>{runner_eff:.1f}건</b>) <br>
                    {"⚠️ <b>참고:</b> 두 조직 간 지국 수(활동 인원)의 편차가 큽니다. 단순 유입량보다 지국당 생산성을 검토할 필요가 있습니다." if abs(winner_jg - runner_jg) / max(winner_jg, 1) > 0.2 else "✅ 두 조직은 지국 수 규모가 유사합니다."}
                </li>
                <li><b>2. 요일 및 골든 타임 현황:</b> <br>
                    - {winner[scope_option]}은 <b>{winner_df['요일'].mode()[0]}요일</b>, {runner_up[scope_option]}은 <b>{runner_df['요일'].mode()[0]}요일</b>에 최대 성과를 기록 중입니다.
                </li>
                <li><b>3. 채널 및 시간대 취약점 (보완 필요):</b> <br>
                    - <b>{winner[scope_option]}</b>: 유입 비중이 가장 낮은 채널은 <b>'{winner_df['유입경로'].value_counts().index[-1]}'</b>이며, 
                      특히 <b>{winner_df.groupby('hour').size().idxmin()}시~{winner_df.groupby('hour').size().idxmin()+2}시</b> 사이의 유입이 매우 취약합니다. <br>
                    - <b>{runner_up[scope_option]}</b>: <b>'{runner_df['유입경로'].value_counts().index[-1]}'</b> 채널 활성화가 시급하며, 
                      <b>{runner_df.groupby('hour').size().idxmin()}시</b> 시간대의 유입 공백을 보완해야 합니다.
                </li>
            </ul>
            <hr>
            <p style='font-weight: bold;'>💡 종합 분석:</p>
            <p>전체적인 소스 유입 규모는 <b>{winner[scope_option]}</b>이 리드하고 있으나, <b>지국 수 대비 유입 효율</b>을 분석한 결과 
            {f"<b>{winner[scope_option] if winner_eff > runner_eff else runner_up[scope_option]}</b>의 조직 가동률이 더 높은 것" if abs(winner_eff - runner_eff) > 0.1 else "두 조직의 가동률은 대등한 수준"}으로 나타납니다. 
            운영 최적화를 위해 각 조직의 취약 시간대인 <b>{winner_df.groupby('hour').size().idxmin()}시</b>와 <b>{runner_df.groupby('hour').size().idxmin()}시</b>를 타겟팅한 채널 보완 전략이 수반되어야 합니다.</p>
            </div>
            """, unsafe_allow_html=True)