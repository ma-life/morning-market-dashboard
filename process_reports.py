"""
증권사 레포트 PDF AI 요약 스크립트
===================================

역할:
    - 폴더 내의 가장 최근에 수정된 PDF 레포트를 탐색합니다.
    - PDF의 텍스트를 추출한 뒤 Claude API(API Key 필요)를 호출하여 
      주요이슈(매크로 요약 2~3개), 종목별 이슈(미국 기업 이슈 3~4개)로 요약합니다.
    - 요약 결과를 'pdf_issues.json'으로 저장하여 대시보드 생성 스크립트와 연동합니다.

사용법:
    1. 로컬 환경변수 또는 같은 폴더 내 `.env` 파일에 ANTHROPIC_API_KEY="your_key" 설정
    2. python process_reports.py
"""

import sys
import os
import json
from pathlib import Path


def load_api_key():
    """환경변수 또는 .env 파일에서 ANTHROPIC_API_KEY 로드"""
    # 1. 시스템 환경변수 확인
    key = os.environ.get("ANTHROPIC_API_KEY")
    if key:
        return key
    
    # 2. .env 파일 확인
    env_path = Path(".env")
    if env_path.exists():
        try:
            for line in env_path.read_text(encoding="utf-8").splitlines():
                if line.strip().startswith("ANTHROPIC_API_KEY="):
                    parts = line.split("=", 1)
                    if len(parts) == 2:
                        return parts[1].strip().strip('"').strip("'")
        except Exception as e:
            print(f".env 파일 읽기 중 오류 발생: {e}")
            
    return None


def get_latest_pdf():
    """폴더 내에서 가장 최근에 수정된 PDF 파일 반환"""
    current_dir = Path(".")
    pdf_files = list(current_dir.glob("*.pdf"))
    if not pdf_files:
        return None
    
    # 수정 시간 기준 정렬 (최신이 가장 뒤)
    pdf_files.sort(key=lambda p: p.stat().st_mtime)
    return pdf_files[-1]


def extract_text_from_pdf(pdf_path, max_pages=15):
    """PDF 파일의 텍스트를 추출 (최대 N페이지 제한으로 토큰 절약)"""
    try:
        import pypdf
    except ImportError:
        print("\n[오류] 'pypdf' 라이브러리가 설치되어 있지 않습니다.")
        print("설치 명령어: pip install pypdf")
        sys.exit(1)
        
    print(f"  텍스트 추출 중: {pdf_path.name} (최대 {max_pages}페이지)")
    try:
        reader = pypdf.PdfReader(pdf_path)
        text = ""
        num_pages = min(len(reader.pages), max_pages)
        for i in range(num_pages):
            page_text = reader.pages[i].extract_text()
            if page_text:
                text += page_text + "\n"
        return text
    except Exception as e:
        print(f"PDF 파싱 실패: {e}")
        sys.exit(1)


def call_claude_api(text, api_key):
    """Claude API를 호출하여 매크로 및 종목 요약 획득"""
    try:
        import requests
    except ImportError:
        print("\n[오류] 'requests' 라이브러리가 설치되어 있지 않습니다.")
        print("설치 명령어: pip install requests")
        sys.exit(1)

    url = "https://api.anthropic.com/v1/messages"
    headers = {
        "x-api-key": api_key,
        "anthropic-version": "2023-06-01",
        "content-type": "application/json"
    }
    
    prompt = (
        "당신은 금융 본부의 리서치 애널리스트입니다. 제공된 증권사 레포트 텍스트를 분석하여, 다음 두 가지 카테고리로 한국어로 요약해 주세요.\n\n"
        "1. 주요 이슈 (전체 글로벌/국내 매크로 시장 요약 2~3개 불릿 포인트)\n"
        "2. 종목별 이슈 (미국 기업 또는 개별 기업 이슈 요약 3~4개 불릿 포인트)\n\n"
        "출력은 반드시 다음과 같은 JSON 형식으로만 반환해 주세요. 다른 설명이나 설명글, 마크다운 기호(```json 제외) 등은 일절 포함하지 마십시오. 오직 JSON 구조만 출력해야 합니다.\n\n"
        "{\n"
        "  \"main_issues\": [\n"
        "    \"매크로 요약 내용 1\",\n"
        "    \"매크로 요약 내용 2\",\n"
        "    \"매크로 요약 내용 3\"\n"
        "  ],\n"
        "  \"stock_issues\": [\n"
        "    \"미국 기업 이슈 요약 내용 1\",\n"
        "    \"미국 기업 이슈 요약 내용 2\",\n"
        "    \"미국 기업 이슈 요약 내용 3\",\n"
        "    \"미국 기업 이슈 요약 내용 4\"\n"
        "  ]\n"
        "}\n\n"
        f"제공된 레포트 텍스트:\n{text}"
    )
    
    data = {
        "model": "claude-3-5-sonnet-20241022",
        "max_tokens": 1200,
        "messages": [
            {"role": "user", "content": prompt}
        ]
    }
    
    print("  Claude API 호출 중 (요약본 생성)...")
    try:
        response = requests.post(url, headers=headers, json=data)
        if response.status_code != 200:
            print(f"\n[오류] API 호출에 실패했습니다 (Status: {response.status_code})")
            print(response.text)
            sys.exit(1)
            
        res_json = response.json()
        raw_text = res_json["content"][0]["text"].strip()
        return raw_text
    except Exception as e:
        print(f"API 호출 중 네트워크 예외 발생: {e}")
        sys.exit(1)


def clean_json_string(s):
    """API 응답 텍스트에서 마크다운 코드 블록 제거"""
    s = s.strip()
    if s.startswith("```json"):
        s = s[7:]
    elif s.startswith("```"):
        s = s[3:]
    if s.endswith("```"):
        s = s[:-3]
    return s.strip()


def main():
    # 1. API 키 확인
    api_key = load_api_key()
    if not api_key:
        print("\n[경고] ANTHROPIC_API_KEY를 찾을 수 없습니다.")
        print("  - 로컬 환경변수에 등록하거나, 폴더 내에 '.env' 파일을 만들어 키를 입력해 주세요.")
        print("  - 예: ANTHROPIC_API_KEY=your_actual_key_here")
        # 실패하더라도 빌드 스크립트가 멈추지 않도록 빈 JSON 구조를 생성해두고 끝냅니다.
        empty_data = {"main_issues": [], "stock_issues": []}
        Path("pdf_issues.json").write_text(json.dumps(empty_data, ensure_ascii=False), encoding="utf-8")
        print("  - 빈 pdf_issues.json 파일을 생성하고 종료합니다.")
        return

    # 2. 최신 PDF 검색
    pdf_path = get_latest_pdf()
    if not pdf_path:
        print("\n[안내] 폴더 내에 요약할 PDF 파일이 존재하지 않습니다.")
        empty_data = {"main_issues": [], "stock_issues": []}
        Path("pdf_issues.json").write_text(json.dumps(empty_data, ensure_ascii=False), encoding="utf-8")
        print("  - 빈 pdf_issues.json 파일을 생성하고 종료합니다.")
        return

    print(f"발견된 최신 레포트: {pdf_path.name}")

    # 3. PDF 텍스트 추출
    pdf_text = extract_text_from_pdf(pdf_path)
    if not pdf_text.strip():
        print("[오류] PDF에서 텍스트를 추출할 수 없습니다. 이미지 기반 스캔 PDF인지 확인해 주세요.")
        sys.exit(1)

    # 4. Claude 요약 API 호출
    api_response = call_claude_api(pdf_text, api_key)
    
    # 5. JSON 정제 및 파싱
    cleaned_json = clean_json_string(api_response)
    try:
        issues_data = json.loads(cleaned_json)
    except Exception as e:
        print(f"[오류] API 응답을 JSON으로 파싱하는 데 실패했습니다: {e}")
        print("원시 응답 데이터:")
        print(api_response)
        sys.exit(1)

    # 6. JSON 파일로 저장
    out_path = Path("pdf_issues.json")
    out_path.write_text(json.dumps(issues_data, ensure_ascii=False, indent=2), encoding="utf-8")
    
    print(f"AI 요약 완료 ➔ {out_path.name} 저장 완료")
    print(f"  - 주요 매크로 이슈: {len(issues_data.get('main_issues', []))}개")
    print(f"  - 종목별 미국 이슈: {len(issues_data.get('stock_issues', []))}개")


if __name__ == "__main__":
    main()
