import streamlit as st
import streamlit.components.v1 as components

# Streamlit 설정: 화면 너비를 최대로 넓히고 탭 제목 지정
st.set_page_config(
    layout="wide",
    page_title="시장 모니터링 대시보드",
    initial_sidebar_state="collapsed"
)

# 좌우 및 상단 마진을 완전히 제거하는 CSS 주입
st.markdown("""
    <style>
        .block-container {
            padding-top: 3.5rem;
            padding-bottom: 0rem;
            padding-left: 0rem;
            padding-right: 0rem;
        }
        iframe {
            display: block;
            border: none;
        }
    </style>
""", unsafe_allow_html=True)

# 로컬에서 컴파일되어 Git으로 Push된 HTML 대시보드 파일 로드
try:
    with open("dashboard_v4_2.html", "r", encoding="utf-8") as f:
        html_content = f.read()
    
    # HTML을 Streamlit 페이지 내에 임베딩 (스크롤 가능하게 높이 넉넉히 설정)
    components.html(html_content, height=1600, scrolling=True)

except FileNotFoundError:
    st.error("대시보드 파일(dashboard_v4_2.html)을 찾을 수 없습니다. 로컬에서 먼저 생성 및 Push를 진행해 주세요.")
