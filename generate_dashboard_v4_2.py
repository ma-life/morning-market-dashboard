"""
시장 모니터링 대시보드 생성 스크립트 v4.2
=========================================

v4.2 추가 기능:
    - v4.1 파일 기반에 차트 인터랙션 보강 (Interactive Chart Tooltip) 추가
    - PDF AI 요약 결과물(pdf_issues.json) 병합 연동 기능 추가

사용법:
    python generate_dashboard_v4_2.py <엑셀파일> [출력HTML]
"""

import sys
import json
from datetime import datetime, timedelta
from pathlib import Path
import openpyxl


# ============================================================
# 설정
# ============================================================

HISTORY_YEARS = 5  # HTML에 임베드할 시계열 기간 (조정 가능)

# ── 로그인 게이트 설정 ──────────────────────────────────────
LOGIN_ENABLED = True
LOGIN_ID = "admin"
LOGIN_PW = "1234"
# ────────────────────────────────────────────────────────────

SPREAD_CATEGORIES = {"Spreads"}
CATEGORY_LABELS = {
    "Spreads": "Spreads (bp)", "Loans": "Loans", "Currency": "Currency",
    "Commodities": "Commodities", "Indices": "Indices",
}
CATEGORY_BADGES = {
    "Spreads": "SP", "Loans": "LN", "Currency": "FX",
    "Commodities": "CM", "Indices": "IX",
}

# Board 지표명 → LAST 시트 ticker 매핑
BOARD_TO_TICKER = {
    "EM Spread": "JPEIGLSP Index",
    "Global HY Spread": "LG30OAS INDEX",
    "US HY Spread": "LF98OAS Index",
    "US Mid-IG Spread": "LUACOAS Index",
    "US Long-IG Spread": "LD07OAS Index",
    "Muni OAS": "BTXMOAS Index",
    "US Loan": "SPBDALB Index",
    "Euro Loan": "SPBDELB Index",
    "달러인덱스": "DXY Curncy",
    "USD/KRW": "KRW REGN Curncy",
    "EUR/USD": "EUR BGN Curncy",
    "USD/CNY": "CNY CURNCY",
    "BRL/KRW": "BRLKRW INDEX",
    "비트코인": "BTC1 Curncy",
    "해상운임 (BDI)": "BDIY Index",
    "WTI": "CLA COMDTY",
    "금": "XAU CURNCY",
    "천연가스": "NG1 COMB Comdty",
    "LNG": "JKL1 COMB Comdty",
    "S&P500": "SPX Index",
    "NASDAQ": "CCMP Index",
    "홍콩H (HSCEI)": "HSCEI Index",
    "코스피": "KOSPI Index",
    "Nikkei225": "NKY Index",
}


# ============================================================
# 엑셀 파싱
# ============================================================

def excel_date_to_dt(v):
    """엑셀 날짜 값(datetime 또는 serial number)을 datetime으로 변환"""
    if isinstance(v, datetime):
        return v
    if isinstance(v, (int, float)):
        return datetime(1899, 12, 30) + timedelta(days=int(v))
    return None


def parse_board_sheet(wb) -> dict:
    """Board 시트 파싱"""
    ws = wb["Board"]
    date_today = ws.cell(3, 4).value
    date_prev = ws.cell(3, 5).value
    date_ytd_base = ws.cell(3, 8).value

    rows = []
    current_category = None
    for r in range(4, 28):
        b_val = ws.cell(r, 2).value
        c_val = ws.cell(r, 3).value
        if b_val and str(b_val).strip():
            current_category = str(b_val).strip()
        if not c_val:
            continue
        rows.append({
            "category": current_category,
            "name": str(c_val).strip(),
            "current": ws.cell(r, 4).value,
        })

    # 이슈 텍스트
    main_issues, stock_issues = [], []
    section = None
    for r in range(29, ws.max_row + 1):
        val = ws.cell(r, 2).value
        if val is None:
            continue
        text = str(val).strip()
        if "[주요 이슈]" in text:
            section = "main"; continue
        if "[종목별 이슈]" in text:
            section = "stock"; continue
        if not text.startswith("-"):
            continue
        cleaned = text.lstrip("- ").strip()
        if len(cleaned) < 5:
            continue
        (main_issues if section == "main" else stock_issues).append(cleaned)

    return {
        "date_today": date_today, "date_prev": date_prev, "date_ytd_base": date_ytd_base,
        "rows": rows, "main_issues": main_issues, "stock_issues": stock_issues,
    }


def parse_last_sheet(wb, years: int, max_date: datetime = None) -> dict:
    """LAST 시트에서 최근 N년 시계열 추출"""
    ws = wb["LAST"]

    # 1) Ticker → 컬럼 번호 매핑 (row 5)
    ticker_cols = {}
    for c in range(3, 27):
        ticker = ws.cell(5, c).value
        if ticker:
            ticker_cols[str(ticker).strip()] = c

    # 2) 컷오프 결정
    anchor = max_date or datetime.now()
    cutoff = anchor - timedelta(days=years * 366)

    series = {ticker: {} for ticker in ticker_cols}

    for r in range(8, ws.max_row + 1):
        date_val = ws.cell(r, 2).value
        dt = excel_date_to_dt(date_val)
        if dt is None:
            continue
        if max_date and dt > max_date:
            continue  # Board의 "today"보다 미래 데이터는 제외
        if dt < cutoff:
            break
        date_str = dt.strftime("%Y-%m-%d")
        for ticker, col in ticker_cols.items():
            v = ws.cell(r, col).value
            if v is not None and isinstance(v, (int, float)):
                series[ticker][date_str] = float(v)

    return series


def detect_multipliers(board_data: dict, last_series: dict) -> dict:
    """오늘 Board값 vs LAST 최신값 비교하여 환산 계수 자동 감지"""
    multipliers = {}
    board_today = {r["name"]: r["current"] for r in board_data["rows"]
                   if isinstance(r["current"], (int, float))}

    if not last_series:
        return multipliers
    all_dates = set()
    for s in last_series.values():
        all_dates.update(s.keys())
    if not all_dates:
        return multipliers
    latest = max(all_dates)

    for board_name, ticker in BOARD_TO_TICKER.items():
        bv = board_today.get(board_name)
        lv = last_series.get(ticker, {}).get(latest)
        if bv is None or lv is None or lv == 0:
            multipliers[board_name] = 1
            continue
        ratio = bv / lv
        for cand in [1, 100, 10, 0.01, 0.1]:
            if 0.95 < ratio / cand < 1.05:
                multipliers[board_name] = cand
                break
        else:
            multipliers[board_name] = 1
    return multipliers


def build_timeseries(board_data, last_series, multipliers) -> dict:
    """프론트엔드용 시계열 데이터 구조 생성"""
    used_tickers = set()
    for row in board_data["rows"]:
        ticker = BOARD_TO_TICKER.get(row["name"])
        if ticker:
            used_tickers.add(ticker)

    all_dates = set()
    for t in used_tickers:
        all_dates.update(last_series.get(t, {}).keys())
    dates_sorted = sorted(all_dates, reverse=True)  # 최신이 앞

    metrics = []
    for row in board_data["rows"]:
        name = row["name"]
        ticker = BOARD_TO_TICKER.get(name)
        mult = multipliers.get(name, 1)
        values = []
        if ticker:
            s = last_series.get(ticker, {})
            for d in dates_sorted:
                v = s.get(d)
                values.append(round(v * mult, 4) if v is not None else None)
        else:
            values = [None] * len(dates_sorted)
        metrics.append({
            "name": name,
            "category": row["category"],
            "isSpread": row["category"] in SPREAD_CATEGORIES,
            "values": values,
        })

    return {
        "dates": dates_sorted,
        "metrics": metrics,
        "ytdBase": board_data["date_ytd_base"].strftime("%Y-%m-%d")
                   if isinstance(board_data["date_ytd_base"], datetime) else None,
        "mainIssues": board_data["main_issues"],
        "stockIssues": board_data["stock_issues"],
        "latestDate": dates_sorted[0] if dates_sorted else None,
    }


# ============================================================
# HTML 생성
# ============================================================

CSS = r"""
* { box-sizing: border-box; }
html, body { margin: 0; padding: 0; }
body {
    font-family: -apple-system, BlinkMacSystemFont, "Apple SD Gothic Neo",
                 "Malgun Gothic", "맑은 고딕", "Segoe UI", Roboto, sans-serif;
    background: #f5f5f4; color: #1f1f1f; font-size: 14px; line-height: 1.55;
    -webkit-font-smoothing: antialiased;
}
.wrap { max-width: 1080px; margin: 0 auto; padding: 24px 20px 80px; }

/* ── Login gate ── */
.login-gate {
    position: fixed; inset: 0; background: #1f2937;
    display: flex; align-items: center; justify-content: center; z-index: 200;
}
.login-box {
    background: #fff; border-radius: 14px; padding: 36px 34px; width: 340px;
    box-shadow: 0 20px 50px rgba(0,0,0,0.3);
    color: #1f1f1f;
}
.login-box h1 { font-size: 20px; font-weight: 600; margin: 0 0 4px; color: #1f1f1f; }
.login-box .sub { font-size: 13px; color: #6b6b6b; margin-bottom: 22px; }
.login-box label { display: block; font-size: 12px; color: #6b6b6b; margin-bottom: 4px; }
.login-box input {
    width: 100%; font: inherit; font-size: 14px; padding: 10px 12px; margin-bottom: 14px;
    border: 1px solid #d6d3d1; border-radius: 8px; background: #fafaf9; color: #1f1f1f;
}
.login-box input:focus { outline: none; border-color: #1d4ed8; background: #fff; }
.login-box button {
    width: 100%; font: inherit; font-size: 14px; font-weight: 500; padding: 11px;
    background: #1f2937; color: #fff; border: none; border-radius: 8px; cursor: pointer;
}
.login-box button:hover { background: #374151; }
.login-error { color: #c53030; font-size: 12px; margin-top: 10px; min-height: 16px; text-align: center; }

/* Header */
.header { padding-bottom: 16px; border-bottom: 1px solid #d6d3d1; margin-bottom: 24px; }
.header h1 { font-size: 22px; font-weight: 600; margin: 0 0 6px; letter-spacing: -0.01em; }
.header .meta { font-size: 13px; color: #6b6b6b; display: flex; gap: 14px; flex-wrap: wrap; align-items: center; }
.header .meta .gen { color: #888; font-size: 12px; }

.controls {
    margin-top: 14px; display: flex; gap: 8px; flex-wrap: wrap; align-items: center;
}
.controls .label { font-size: 12px; color: #6b6b6b; }
.controls input[type=date] {
    font: inherit; font-size: 13px; padding: 5px 9px;
    border: 1px solid #d6d3d1; background: #fff; border-radius: 6px;
    color: #1f1f1f;
}
.controls button {
    font: inherit; font-size: 12px; padding: 6px 12px;
    border: 1px solid #d6d3d1; background: #fff; border-radius: 6px;
    cursor: pointer; color: #1f1f1f;
}
.controls button:hover { background: #f5f5f4; }
.controls button.reset { color: #1d4ed8; }
.history-mode {
    display: inline-block; padding: 3px 8px; font-size: 11px;
    background: #fef3c7; color: #92400e; border-radius: 4px; font-weight: 500;
}

/* Category */
.category { margin-bottom: 28px; }
.cat-head { display: flex; align-items: center; gap: 10px; margin-bottom: 10px; }
.cat-badge {
    display: inline-flex; align-items: center; justify-content: center;
    width: 26px; height: 22px; background: #1f2937; color: #fff;
    font-size: 10px; font-weight: 600; letter-spacing: 0.05em; border-radius: 4px;
}
.cat-title { font-size: 14px; font-weight: 600; color: #404040; }
.cat-count { font-size: 12px; color: #888; }

/* Table */
table { width: 100%; border-collapse: collapse; background: #fff; border-radius: 8px; overflow: hidden; }
table th, table td { padding: 9px 12px; font-size: 13px; }
table th {
    background: #fafaf9; color: #6b6b6b; font-weight: 500;
    text-align: right; border-bottom: 1px solid #e7e5e4;
    font-size: 11px; letter-spacing: 0.02em;
}
table th.text-left { text-align: left; }
table td { border-top: 1px solid #f0efed; font-feature-settings: "tnum"; }
table td:first-child { font-weight: 500; }
table td.num { text-align: right; }
table tr.metric-row { cursor: pointer; transition: background 0.1s; }
table tr.metric-row:hover td { background: #f0f9ff; }
.up { color: #c53030; }     /* 빨강 = 상승 */
.down { color: #1d4ed8; }   /* 파랑 = 하락 */
.neutral { color: #6b6b6b; }
.spark { width: 100px; height: 22px; vertical-align: middle; }
.spark-cell { text-align: center; padding: 4px 8px !important; width: 110px; }

/* Issue cards */
.issues { display: grid; grid-template-columns: 1fr 1fr; gap: 16px; margin-top: 28px; }
.issues.hidden { display: none; }
.issues-note {
    margin-top: 28px; padding: 14px 16px;
    background: #fef3c7; border-radius: 8px;
    font-size: 13px; color: #92400e;
}
.issues-note.hidden { display: none; }
.issue-card { background: #fff; border-radius: 10px; padding: 16px 18px; border: 1px solid #e7e5e4; }
.issue-card h2 {
    font-size: 13px; font-weight: 600; margin: 0 0 10px;
    display: flex; align-items: center; gap: 8px; color: #404040;
}
.issue-card.main h2 .dot { background: #1d4ed8; }
.issue-card.stock h2 .dot { background: #c53030; }
.dot { display: inline-block; width: 7px; height: 7px; border-radius: 50%; }
.issue-card ul { list-style: none; margin: 0; padding: 0; }
.issue-card li {
    padding: 6px 0 6px 14px; position: relative; font-size: 13px; line-height: 1.6;
    color: #2b2b2b; border-top: 1px dashed #f0efed;
}
.issue-card li:first-child { border-top: none; padding-top: 0; }
.issue-card li:before { content: "·"; position: absolute; left: 4px; color: #888; }
.issue-card li .badge-ai {
    display: inline-block; background: #e0f2fe; color: #0369a1; 
    font-size: 10px; font-weight: bold; padding: 1px 4px; border-radius: 3px; margin-right: 4px;
}

/* Modal */
.modal-backdrop {
    display: none; position: fixed; inset: 0; background: rgba(0,0,0,0.5);
    z-index: 100; align-items: center; justify-content: center;
}
.modal-backdrop.open { display: flex; }
.modal {
    background: #fff; border-radius: 12px; padding: 24px;
    width: min(900px, calc(100vw - 32px));
    max-height: calc(100vh - 32px); overflow: auto;
}
.modal h2 { margin: 0 0 4px; font-size: 18px; font-weight: 600; }
.modal .modal-sub { font-size: 13px; color: #6b6b6b; margin-bottom: 16px; }
.modal-close {
    float: right; cursor: pointer; border: none; background: none;
    font-size: 22px; color: #888; padding: 0; line-height: 1;
}

/* Tooltip & Interactive Charting */
.chart-container {
    position: relative; width: 100%; height: 360px;
    background: rgba(0,0,0,0.01); border-radius: 8px;
    border: 1px solid #e7e5e4; overflow: hidden;
}
.modal-chart { width: 100%; height: 100%; display: block; }
.chart-tooltip {
    position: absolute;
    background: rgba(31, 41, 55, 0.95);
    backdrop-filter: blur(8px);
    -webkit-backdrop-filter: blur(8px);
    border: 1px solid rgba(255, 255, 255, 0.15);
    border-radius: 8px;
    padding: 8px 12px;
    color: #fff;
    font-size: 12px;
    pointer-events: none;
    box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.3);
    z-index: 1000;
    display: none;
    min-width: 140px;
}
.chart-tooltip .date {
    color: #9ca3af; font-size: 10px; margin-bottom: 4px; font-weight: 500;
}
.chart-tooltip .val {
    font-size: 14px; font-weight: 700; font-family: -apple-system, BlinkMacSystemFont, monospace; color: #fff;
}
.chart-tooltip .change {
    font-size: 11px; margin-top: 2px; font-weight: 500;
}

.modal-stats {
    display: grid; grid-template-columns: repeat(auto-fit, minmax(120px, 1fr));
    gap: 10px; margin-top: 16px;
}
.modal-stat { background: #fafaf9; padding: 10px 12px; border-radius: 6px; }
.modal-stat .l { font-size: 11px; color: #6b6b6b; }
.modal-stat .v { font-size: 14px; font-weight: 600; margin-top: 2px; font-feature-settings: "tnum"; }

/* Footer */
.footer { margin-top: 36px; padding-top: 14px; border-top: 1px solid #e7e5e4;
          font-size: 11px; color: #888; text-align: center; }

/* Mobile */
@media (max-width: 760px) {
    .wrap { padding: 16px 14px 60px; }
    .header h1 { font-size: 18px; }
    .issues { grid-template-columns: 1fr; }
    table th, table td { padding: 7px 8px; font-size: 12px; }
    .col-hide-mobile { display: none; }
    .spark { width: 60px; }
    .spark-cell { width: 70px; padding: 4px !important; }
}

/* Print */
@media print {
    body { background: #fff; }
    .wrap { padding: 12px; max-width: none; }
    .controls, .modal-backdrop, .logout-btn { display: none !important; }
    .category, .issue-card { break-inside: avoid; }
}
"""

JS = r"""
const STATE = { dateIdx: 0, activeMetric: null };
let activeChartData = []; // mousemove 시 탐색 대상 데이터 좌표 리스트

// 세션 저장
function safeSession(action, key, val) {
    try {
        if (action === 'get') return sessionStorage.getItem(key);
        if (action === 'set') sessionStorage.setItem(key, val);
    } catch (e) { return null; }
}

// ============================================================
// 로그인 게이트
// ============================================================
function initLogin() {
    if (!CONFIG.loginEnabled) { showApp(); return; }
    if (safeSession('get', 'mm_auth_v4_2') === '1') { showApp(); return; }

    const gate = document.getElementById('login-gate');
    gate.style.display = 'flex';
    
    const submit = () => {
        const id = document.getElementById('login-id').value.trim();
        const pw = document.getElementById('login-pw').value;
        if (id === CONFIG.loginId && pw === CONFIG.loginPw) {
            safeSession('set', 'mm_auth_v4_2', '1');
            showApp();
        } else {
            document.getElementById('login-error').textContent = 'ID 또는 비밀번호가 올바르지 않습니다.';
        }
    };
    
    document.getElementById('login-submit').addEventListener('click', submit);
    document.getElementById('login-pw').addEventListener('keydown', e => { if (e.key === 'Enter') submit(); });
    document.getElementById('login-id').addEventListener('keydown', e => { if (e.key === 'Enter') document.getElementById('login-pw').focus(); });
    document.getElementById('login-id').focus();
}

function logout() {
    try { sessionStorage.removeItem('mm_auth_v4_2'); } catch (e) {}
    location.reload();
}

function showApp() {
    const gate = document.getElementById('login-gate');
    if (gate) gate.style.display = 'none';
    document.getElementById('app').style.display = 'block';
    initApp();
}

// ============================================================
// 앱 초기화 로직
// ============================================================
function initApp() {
    if (!DATA.dates || DATA.dates.length === 0) return;
    const picker = document.getElementById('date-picker');
    picker.min = DATA.dates[DATA.dates.length - 1];
    picker.max = DATA.dates[0];
    picker.value = DATA.dates[0];
    picker.addEventListener('change', e => {
        const idx = DATA.dates.indexOf(e.target.value);
        if (idx >= 0) {
            STATE.dateIdx = idx;
            render();
        } else {
            const target = e.target.value;
            let closest = 0;
            for (let i = 0; i < DATA.dates.length; i++) {
                if (DATA.dates[i] <= target) { closest = i; break; }
            }
            STATE.dateIdx = closest;
            picker.value = DATA.dates[closest];
            render();
        }
    });
    document.getElementById('reset-btn').addEventListener('click', () => {
        STATE.dateIdx = 0;
        picker.value = DATA.dates[0];
        render();
    });
    document.getElementById('modal-backdrop').addEventListener('click', e => {
        if (e.target.id === 'modal-backdrop') closeModal();
    });
    
    // 차트 마우스 리스너 바인딩
    const chartSvg = document.getElementById('modal-chart');
    chartSvg.addEventListener('mousemove', handleChartMouseMove);
    chartSvg.addEventListener('mouseleave', handleChartMouseLeave);
    
    render();
}

function render() {
    renderHeader();
    renderTable();
    renderIssuesVisibility();
}

function renderHeader() {
    const d = DATA.dates[STATE.dateIdx];
    document.getElementById('display-date').textContent = formatDate(d);
    const isHistory = STATE.dateIdx > 0;
    document.getElementById('history-badge').style.display = isHistory ? 'inline-block' : 'none';
    document.getElementById('reset-btn').style.display = isHistory ? 'inline-block' : 'none';
}

function renderTable() {
    const container = document.getElementById('tables-container');
    container.innerHTML = '';
    // 카테고리별 그룹화
    const byCategory = {};
    const catOrder = [];
    for (const m of DATA.metrics) {
        const cat = m.category || '기타';
        if (!byCategory[cat]) { byCategory[cat] = []; catOrder.push(cat); }
        byCategory[cat].push(m);
    }
    for (const cat of catOrder) {
        container.appendChild(buildCategoryBlock(cat, byCategory[cat]));
    }
}

const CATEGORY_LABELS = {
    'Spreads': 'Spreads (bp)', 'Loans': 'Loans', 'Currency': 'Currency',
    'Commodities': 'Commodities', 'Indices': 'Indices',
};
const CATEGORY_BADGES = {
    'Spreads': 'SP', 'Loans': 'LN', 'Currency': 'FX',
    'Commodities': 'CM', 'Indices': 'IX',
};

function buildCategoryBlock(category, metrics) {
    const block = document.createElement('div');
    block.className = 'category';
    block.innerHTML = `
        <div class="cat-head">
            <span class="cat-badge">${CATEGORY_BADGES[category] || '··'}</span>
            <span class="cat-title">${CATEGORY_LABELS[category] || category}</span>
            <span class="cat-count">${metrics.length}개 지표</span>
        </div>
        <table>
            <thead>
                <tr>
                    <th class="text-left">지표</th>
                    <th>현재</th>
                    <th class="col-hide-mobile">전영업일</th>
                    <th>전일대비</th>
                    <th>YTD</th>
                    <th>최근 60일</th>
                </tr>
            </thead>
            <tbody></tbody>
        </table>
    `;
    const tbody = block.querySelector('tbody');
    for (const m of metrics) {
        tbody.appendChild(buildMetricRow(m));
    }
    return block;
}

function buildMetricRow(metric) {
    const tr = document.createElement('tr');
    tr.className = 'metric-row';
    tr.title = '클릭하여 5년 차트 보기';
    
    const idx = STATE.dateIdx;
    const cur = metric.values[idx];
    const prev = metric.values[idx + 1];
    
    let dod = null;
    if (cur !== null && prev !== null) {
        dod = metric.isSpread ? (cur - prev) : ((cur - prev) / prev * 100);
    }
    
    const selectedDate = DATA.dates[idx];
    const selectedYear = parseInt(selectedDate.substring(0, 4));
    const prevYearEnd = (selectedYear - 1) + '-12-31';
    let ytdBaseIdx = -1;
    for (let i = idx; i < DATA.dates.length; i++) {
        if (DATA.dates[i] <= prevYearEnd) { ytdBaseIdx = i; break; }
    }
    const ytdBase = ytdBaseIdx >= 0 ? metric.values[ytdBaseIdx] : null;
    let ytd = null;
    if (cur !== null && ytdBase !== null && ytdBase !== 0) {
        ytd = metric.isSpread ? (cur - ytdBase) : ((cur - ytdBase) / ytdBase * 100);
    }
    
    const dodText = fmtChange(dod, metric.isSpread);
    const ytdText = fmtChange(ytd, metric.isSpread);
    const curText = fmtValue(cur, metric);
    const prevText = fmtValue(prev, metric);
    
    tr.innerHTML = `
        <td>${escapeHTML(metric.name)}</td>
        <td class="num">${curText}</td>
        <td class="num col-hide-mobile">${prevText}</td>
        <td class="num ${dodText.cls}">${dodText.text}</td>
        <td class="num ${ytdText.cls}">${ytdText.text}</td>
        <td class="spark-cell">${buildSparkline(metric)}</td>
    `;
    tr.addEventListener('click', () => openModal(metric));
    return tr;
}

function buildSparkline(metric) {
    const N = 60;
    const start = Math.max(0, STATE.dateIdx);
    const end = Math.min(DATA.dates.length, start + N);
    const slice = metric.values.slice(start, end).filter(v => v !== null);
    if (slice.length < 2) return '<svg class="spark"></svg>';
    
    const W = 100, H = 22, P = 2;
    const min = Math.min(...slice), max = Math.max(...slice);
    const range = max - min || 1;
    
    const reversed = slice.slice().reverse();
    const points = reversed.map((v, i) => {
        const x = P + (i / (reversed.length - 1)) * (W - 2 * P);
        const y = H - P - ((v - min) / range) * (H - 2 * P);
        return `${x.toFixed(1)},${y.toFixed(1)}`;
    }).join(' ');
    
    const trend = reversed[reversed.length - 1] - reversed[0];
    const color = trend > 0 ? '#c53030' : trend < 0 ? '#1d4ed8' : '#888';
    
    return `<svg class="spark" viewBox="0 0 ${W} ${H}"><polyline fill="none" stroke="${color}" stroke-width="1.2" points="${points}"/></svg>`;
}

function renderIssuesVisibility() {
    const isHistory = STATE.dateIdx > 0;
    document.getElementById('issues-block').classList.toggle('hidden', isHistory);
    document.getElementById('issues-note').classList.toggle('hidden', !isHistory);
}

// ============================================================
// Modal - 확대 차트
// ============================================================
function openModal(metric) {
    STATE.activeMetric = metric;
    
    document.getElementById('modal-title').textContent = metric.name;
    document.getElementById('modal-category').textContent = CATEGORY_LABELS[metric.category] || metric.category;
    
    const valid = metric.values.filter(v => v !== null);
    const min = Math.min(...valid), max = Math.max(...valid);
    const cur = metric.values[STATE.dateIdx];
    const avg = valid.reduce((a, b) => a + b, 0) / valid.length;
    
    document.getElementById('stat-current').textContent = fmtValue(cur, metric);
    document.getElementById('stat-min').textContent = fmtValue(min, metric);
    document.getElementById('stat-max').textContent = fmtValue(max, metric);
    document.getElementById('stat-avg').textContent = fmtValue(avg, metric);
    
    drawFullChart(metric);
    document.getElementById('modal-backdrop').classList.add('open');
}

function closeModal() {
    document.getElementById('modal-backdrop').classList.remove('open');
    STATE.activeMetric = null;
    handleChartMouseLeave();
}

function drawFullChart(metric) {
    const svg = document.getElementById('modal-chart');
    const W = svg.clientWidth || 800, H = 360;
    const PAD_L = 60, PAD_R = 20, PAD_T = 20, PAD_B = 40;
    const innerW = W - PAD_L - PAD_R, innerH = H - PAD_T - PAD_B;
    
    const dates = DATA.dates.slice().reverse();
    const values = metric.values.slice().reverse();
    
    const valid = [];
    for (let i = 0; i < values.length; i++) {
        if (values[i] !== null) valid.push({ d: dates[i], v: values[i], i: i });
    }
    if (valid.length < 2) {
        svg.innerHTML = '<text x="50%" y="50%" text-anchor="middle" fill="#888">데이터 부족</text>';
        return;
    }
    
    const vals = valid.map(p => p.v);
    let min = Math.min(...vals), max = Math.max(...vals);
    const pad = (max - min) * 0.05;
    min -= pad; max += pad;
    const range = max - min || 1;
    
    const selectedDate = DATA.dates[STATE.dateIdx];
    const selIdx = valid.findIndex(p => p.d === selectedDate);
    
    const xScale = i => PAD_L + (i / (valid.length - 1)) * innerW;
    const yScale = v => PAD_T + innerH - ((v - min) / range) * innerH;
    
    activeChartData = valid.map((p, i) => ({
        d: p.d,
        v: p.v,
        x: xScale(i),
        y: yScale(p.v),
        prevV: i > 0 ? valid[i - 1].v : null
    }));
    
    const parts = [];
    
    for (let k = 0; k <= 4; k++) {
        const yv = min + (k / 4) * range;
        const y = yScale(yv);
        parts.push(`<line x1="${PAD_L}" y1="${y}" x2="${W - PAD_R}" y2="${y}" stroke="#e7e5e4" stroke-width="0.5"/>`);
        parts.push(`<text x="${PAD_L - 6}" y="${y + 4}" text-anchor="end" font-size="10" fill="#888" font-family="sans-serif">${fmtAxis(yv, metric)}</text>`);
    }
    
    for (let k = 0; k <= 4; k++) {
        const idx = Math.floor((valid.length - 1) * k / 4);
        const x = xScale(idx);
        const d = valid[idx].d;
        parts.push(`<text x="${x}" y="${H - PAD_B + 16}" text-anchor="middle" font-size="10" fill="#888" font-family="sans-serif">${d.substring(2, 7)}</text>`);
    }
    
    if (selIdx >= 0) {
        const x = xScale(selIdx);
        parts.push(`<line x1="${x}" y1="${PAD_T}" x2="${x}" y2="${H - PAD_B}" stroke="#c53030" stroke-width="1" stroke-dasharray="3,3"/>`);
    }
    
    const points = valid.map((p, i) => `${xScale(i)},${yScale(p.v)}`).join(' ');
    parts.push(`<polyline fill="none" stroke="#1f2937" stroke-width="1.3" points="${points}"/>`);
    
    if (selIdx >= 0) {
        const p = valid[selIdx];
        parts.push(`<circle cx="${xScale(selIdx)}" cy="${yScale(p.v)}" r="4" fill="#c53030"/>`);
    }
    
    parts.push(`<line id="hover-guide-line" x1="0" y1="${PAD_T}" x2="0" y2="${H - PAD_B}" stroke="rgba(29, 78, 216, 0.4)" stroke-width="1.2" stroke-dasharray="3,3" style="display:none;"/>`);
    parts.push(`<circle id="hover-guide-point" cx="0" cy="0" r="4.5" fill="#1d4ed8" stroke="#fff" stroke-width="1" style="display:none;"/>`);
    
    svg.innerHTML = parts.join('');
}

function handleChartMouseMove(e) {
    if (!STATE.activeMetric || activeChartData.length < 2) return;
    const svg = document.getElementById('modal-chart');
    const rect = svg.getBoundingClientRect();
    
    const mouseX = e.clientX - rect.left;
    
    let closestIdx = 0;
    let minDiff = Infinity;
    for (let i = 0; i < activeChartData.length; i++) {
        const diff = Math.abs(activeChartData[i].x - mouseX);
        if (diff < minDiff) {
            minDiff = diff;
            closestIdx = i;
        }
    }
    
    const pt = activeChartData[closestIdx];
    
    const line = document.getElementById('hover-guide-line');
    const circle = document.getElementById('hover-guide-point');
    if (line && circle) {
        line.setAttribute('x1', pt.x.toFixed(1));
        line.setAttribute('x2', pt.x.toFixed(1));
        line.style.display = 'block';
        
        circle.setAttribute('cx', pt.x.toFixed(1));
        circle.setAttribute('cy', pt.y.toFixed(1));
        circle.style.display = 'block';
    }
    
    const tooltip = document.getElementById('chart-tooltip');
    tooltip.style.display = 'block';
    
    let diffText = '';
    if (pt.prevV !== null) {
        const diff = STATE.activeMetric.isSpread ? (pt.v - pt.prevV) : ((pt.v - pt.prevV) / pt.prevV * 100);
        const fmted = fmtChange(diff, STATE.activeMetric.isSpread);
        diffText = `<span class="${fmted.cls}">${fmted.text}</span>`;
    } else {
        diffText = `<span class="neutral">-</span>`;
    }
    
    tooltip.innerHTML = `
        <div class="date">${formatDate(pt.d)}</div>
        <div class="val">${fmtValue(pt.v, STATE.activeMetric)}</div>
        <div class="change">${diffText}</div>
    `;
    
    const tooltipW = tooltip.offsetWidth;
    const tooltipH = tooltip.offsetHeight;
    
    let leftPos = pt.x - tooltipW / 2;
    let topPos = pt.y - tooltipH - 10;
    
    if (leftPos < 10) leftPos = 10;
    if (leftPos + tooltipW > rect.width - 10) leftPos = rect.width - tooltipW - 10;
    if (topPos < 10) topPos = pt.y + 10;
    
    tooltip.style.left = `${leftPos}px`;
    tooltip.style.top = `${topPos}px`;
}

function handleChartMouseLeave() {
    const line = document.getElementById('hover-guide-line');
    const circle = document.getElementById('hover-guide-point');
    if (line && circle) {
        line.style.display = 'none';
        circle.style.display = 'none';
    }
    const tooltip = document.getElementById('chart-tooltip');
    if (tooltip) tooltip.style.display = 'none';
}

function fmtValue(v, metric) {
    if (v === null || v === undefined) return '-';
    if (metric.isSpread) {
        return v >= 100 ? Math.round(v).toLocaleString() : v.toFixed(2);
    }
    if (metric.category === 'Loans') return v.toFixed(4);
    if (Math.abs(v) >= 1000) return v.toLocaleString(undefined, { maximumFractionDigits: 2 });
    if (Math.abs(v) >= 10) return v.toFixed(2);
    return v.toFixed(4);
}

function fmtAxis(v, metric) {
    if (metric.isSpread) return Math.round(v).toString();
    if (Math.abs(v) >= 1000) return Math.round(v).toLocaleString();
    return v.toFixed(2);
}

function fmtChange(v, isSpread) {
    if (v === null || v === undefined) return { text: '-', cls: 'neutral' };
    const cls = v > 0 ? 'up' : v < 0 ? 'down' : 'neutral';
    const sign = v > 0 ? '+' : '';
    if (isSpread) {
        return { text: sign + v.toFixed(2), cls: cls };
    }
    return { text: sign + v.toFixed(2) + '%', cls: cls };
}

function formatDate(d) {
    if (!d) return '';
    const dt = new Date(d + 'T00:00:00');
    const wd = '월화수목금토일'[(dt.getDay() + 6) % 7];
    return `${d} (${wd})`;
}

function escapeHTML(s) {
    return String(s).replace(/[&<>"']/g, c => ({
        '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
    }[c]));
}

function printPage() { window.print(); }
function copyLink() {
    navigator.clipboard.writeText(location.href).then(
        () => alert('링크가 복사되었습니다'),
        () => alert('복사 실패 - 주소창에서 직접 복사해 주세요')
    );
}

document.addEventListener('DOMContentLoaded', initLogin);
"""


def generate_html(board_data, timeseries) -> str:
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M")
    
    # [AI] 말머리가 붙은 것은 뱃지를 노출하도록 렌더링 변경
    def format_issue_li(issue):
        if issue.startswith("[AI]"):
            content = issue[4:].strip()
            return f"<li><span class='badge-ai'>AI 요약</span>{content}</li>"
        return f"<li>{issue}</li>"

    main_lis = "".join(format_issue_li(i) for i in board_data["main_issues"])
    stock_lis = "".join(format_issue_li(i) for i in board_data["stock_issues"])

    data_json = json.dumps({
        "dates": timeseries["dates"],
        "metrics": timeseries["metrics"],
        "ytdBase": timeseries["ytdBase"],
        "latestDate": timeseries["latestDate"],
    }, ensure_ascii=False, separators=(",", ":"))

    config_json = json.dumps({
        "loginEnabled": LOGIN_ENABLED,
        "loginId": LOGIN_ID,
        "loginPw": LOGIN_PW,
        "historyYears": HISTORY_YEARS,
    }, ensure_ascii=False)

    return f"""<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>시장 모니터링 대시보드 v4.2</title>
<style>{CSS}</style>
</head>
<body>

<!-- 로그인 게이트 -->
<div class="login-gate" id="login-gate" style="display:none;">
    <div class="login-box">
        <h1>시장 모니터링</h1>
        <div class="sub">본부 내부 공유용 대시보드 (v4.2)</div>
        <label>ID</label>
        <input type="text" id="login-id" autocomplete="username">
        <label>비밀번호</label>
        <input type="password" id="login-pw" autocomplete="current-password">
        <button id="login-submit">로그인</button>
        <div class="login-error" id="login-error"></div>
    </div>
</div>

<!-- 메인 앱 -->
<div id="app" style="display:none;">
<div class="wrap">
    <header class="header">
        <h1>시장 모니터링</h1>
        <div class="meta">
            <span>기준일: <strong id="display-date">-</strong></span>
            <span id="history-badge" class="history-mode" style="display:none;">과거 일자 조회 중</span>
            <span class="gen">생성: {now_str}</span>
            <button class="logout-btn" onclick="logout()" style="margin-left: 10px; font: inherit; font-size: 12px; color: #6b6b6b; background: none; border: none; cursor: pointer; text-decoration: underline;">로그아웃</button>
        </div>
        <div class="controls">
            <span class="label">일자 선택:</span>
            <input type="date" id="date-picker">
            <button id="reset-btn" class="reset" style="display:none;">오늘로</button>
            <button onclick="printPage()">인쇄</button>
            <button onclick="copyLink()">링크 복사</button>
        </div>
    </header>

    <div id="tables-container"></div>

    <div class="issues-note hidden" id="issues-note">
        과거 일자 조회 모드입니다. 주요 이슈·종목별 이슈는 매일 작성되는 당일 시점 정보이므로 과거 일자에서는 표시되지 않습니다.
    </div>

    <div class="issues" id="issues-block">
        <div class="issue-card main">
            <h2><span class="dot"></span>주요 이슈</h2>
            <ul>{main_lis}</ul>
        </div>
        <div class="issue-card stock">
            <h2><span class="dot"></span>종목별 이슈</h2>
            <ul>{stock_lis}</ul>
        </div>
    </div>

    <div class="footer">
        본 자료는 본부 내부 공유용입니다. · 시계열: 최근 {HISTORY_YEARS}년 · 자동 생성됨
    </div>
</div>
</div>

<!-- Modal -->
<div class="modal-backdrop" id="modal-backdrop">
    <div class="modal">
        <button class="modal-close" onclick="closeModal()">×</button>
        <h2 id="modal-title">-</h2>
        <div class="modal-sub" id="modal-category">-</div>
        
        <!-- 차트 컨테이너 및 툴팁 -->
        <div class="chart-container">
            <svg id="modal-chart" class="modal-chart" viewBox="0 0 800 360" preserveAspectRatio="none"></svg>
            <div id="chart-tooltip" class="chart-tooltip"></div>
        </div>
        
        <div class="modal-stats">
            <div class="modal-stat"><div class="l">선택 일자</div><div class="v" id="stat-current">-</div></div>
            <div class="modal-stat"><div class="l">5년 최저</div><div class="v" id="stat-min">-</div></div>
            <div class="modal-stat"><div class="l">5년 최고</div><div class="v" id="stat-max">-</div></div>
            <div class="modal-stat"><div class="l">5년 평균</div><div class="v" id="stat-avg">-</div></div>
        </div>
    </div>
</div>

<script>
const DATA = {data_json};
const CONFIG = {config_json};
{JS}
</script>
</body>
</html>
"""


def main():
    if len(sys.argv) < 2:
        print("사용법: python generate_dashboard_v4_2.py <엑셀파일> [출력HTML]")
        sys.exit(1)

    xlsx_path = sys.argv[1]
    out_path = sys.argv[2] if len(sys.argv) > 2 else "dashboard_v4_2.html"

    if not Path(xlsx_path).exists():
        print(f"파일을 찾을 수 없습니다: {xlsx_path}")
        sys.exit(1)

    print(f"읽는 중: {xlsx_path}")
    wb = openpyxl.load_workbook(xlsx_path, data_only=True, keep_vba=False)

    print("  Board 시트 파싱...")
    board = parse_board_sheet(wb)
    print(f"    엑셀 원본 지표 {len(board['rows'])}개, 주요이슈 {len(board['main_issues'])}개, "
          f"종목이슈 {len(board['stock_issues'])}개")

    # ── AI PDF 요약본 병합 연동 ────────────────────────────────
    pdf_issues_path = Path("pdf_issues.json")
    if pdf_issues_path.exists():
        print("  [AI 연동] pdf_issues.json 파일 발견! 대시보드 이슈 병합 중...")
        try:
            pdf_data = json.loads(pdf_issues_path.read_text(encoding="utf-8"))
            pdf_main = pdf_data.get("main_issues", [])
            pdf_stock = pdf_data.get("stock_issues", [])
            
            # 중복 검사하며 병합
            for issue in pdf_main:
                issue_stripped = issue.strip()
                if issue_stripped and not any(issue_stripped in x for x in board["main_issues"]):
                    board["main_issues"].append(f"[AI] {issue_stripped}")
            for issue in pdf_stock:
                issue_stripped = issue.strip()
                if issue_stripped and not any(issue_stripped in x for x in board["stock_issues"]):
                    board["stock_issues"].append(f"[AI] {issue_stripped}")
                    
            print(f"    병합 후 최종 주요이슈: {len(board['main_issues'])}개, 종목이슈: {len(board['stock_issues'])}개")
        except Exception as e:
            print(f"  [경고] AI 요약 파일 로드 실패: {e}")
    # ──────────────────────────────────────────────────────────

    print(f"  LAST 시트 파싱 (최근 {HISTORY_YEARS}년)...")
    board_today_dt = board["date_today"] if isinstance(board["date_today"], datetime) else None
    last = parse_last_sheet(wb, HISTORY_YEARS, max_date=board_today_dt)
    n_dates = len(set().union(*[s.keys() for s in last.values()])) if last else 0
    print(f"    {len(last)}개 ticker, {n_dates}영업일 (cap: {board_today_dt.date() if board_today_dt else 'none'})")

    print("  단위 환산 계수 자동 감지...")
    mults = detect_multipliers(board, last)
    spread_mults = {k: v for k, v in mults.items() if v != 1}
    print(f"    ×100 환산 대상: {list(spread_mults.keys())}")

    print("  시계열 데이터 구조화...")
    ts = build_timeseries(board, last, mults)

    html = generate_html(board, ts)
    Path(out_path).write_text(html, encoding="utf-8")
    size_kb = len(html) / 1024
    print(f"생성 완료: {out_path} ({size_kb:.0f} KB)")
    if LOGIN_ENABLED:
        print(f"  로그인: ID='{LOGIN_ID}' PW='{LOGIN_PW}' (보안 게이트 및 인터랙티브 툴팁 설정 완료)")


if __name__ == "__main__":
    main()
