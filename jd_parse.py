"""
공고 분석 에이전트 v2 — raw_data 파싱 → jd_analysis 자동 추가
변경사항: raw_data에서 회사명·직무명 컬럼 제거, LLM이 원문에서 자동 추출
컬럼: A(공고ID) B(트랙) C(공고링크) D(원문텍스트) E(수집일자) F(상태)
사용법: python3 jd_parse_v2.py
"""

import os
import json
import openpyxl
import anthropic

# ================================================================
# 설정 — 필요 시 여기만 수정
# ================================================================
EXCEL_PATH = "/Users/maczoo/Documents/Vault/Attached/jd_analysis.xlsm"
API_KEY    = os.environ.get("ANTHROPIC_API_KEY", "")
MODEL      = "claude-sonnet-4-6"

# ================================================================
# 프롬프트
# ================================================================
SYSTEM_PROMPT = """당신은 채용공고 파싱 전문가입니다.
채용공고 원문을 읽고 아래 JSON 형식으로 정확히 파싱해서 반환하세요.
반드시 JSON만 반환하고, 설명이나 마크다운 코드블록은 절대 포함하지 마세요.

반환 형식:
{
  "회사명": "회사명 (원문에서 추출, 없으면 빈 문자열)",
  "직무명": "직무명/포지션명 (원문에서 추출)",
  "업무": ["업무 항목1", "업무 항목2", ...],
  "자격요건": [
    {"구분": "필수", "내용": "필수 항목 내용"},
    {"구분": "우대", "내용": "우대 항목 내용"},
    ...
  ],
  "연차": "연차 정보 (예: 3년 이상, 신입, 무관 등. 명시 없으면 빈 문자열)",
  "근무형태": "정규직/계약직/인턴 등 (명시 없으면 빈 문자열)",
  "근무지": "근무 위치 (명시 없으면 빈 문자열)"
}

규칙:
- 회사명과 직무명은 원문 첫 부분이나 제목에서 추출하세요.
- 업무(주요 업무/담당 업무)와 자격요건(필수/우대)을 명확히 구분하세요.
- 자격요건이 필수/우대로 구분되지 않은 경우 문맥상 판단해서 분류하세요.
- 각 항목은 원문의 의미를 유지하되 한 항목당 하나의 내용만 담아주세요.
- 연차는 자격요건 전체에 공통으로 적용되는 값만 적고, 항목별로 다르면 빈 문자열로 두세요.
"""

# ================================================================
# raw_data 컬럼 위치 (변경 시 여기만 수정)
# ================================================================
COL_RAW = {
    "job_id":    1,  # A: 공고ID
    "track":     2,  # B: 트랙
    "link":      3,  # C: 공고 링크
    "raw_text":  4,  # D: 원문 텍스트
    "collected": 5,  # E: 수집일자
    "status":    6,  # F: 상태
}

# jd_analysis 컬럼 위치
COL_JD = {
    "job_id":    1,   # A: 공고ID
    "company":   2,   # B: 회사명
    "track":     3,   # C: 트랙
    "job_title": 4,   # D: 직무명
    # E열: 빈 열
    "업무":      6,   # F: 업무
    "구분":      7,   # G: 필수/우대
    "내용":      8,   # H: 내용
    "연차":      9,   # I: 연차
    "link":      10,  # J: 공고 링크
    "근무형태":  11,  # K: 근무 형태
    "근무지":    12,  # L: 근무지
    "collected": 14,  # N: 수집일자
}


def parse_jd(client, raw_text):
    """원문 텍스트를 LLM으로 파싱"""
    response = client.messages.create(
        model=MODEL,
        max_tokens=2000,
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": f"다음 채용공고 원문을 파싱해주세요.\n\n{raw_text}"}]
    )
    raw_json = response.content[0].text.strip()
    return json.loads(raw_json)


def get_last_row(ws):
    """실제 데이터가 있는 마지막 행 번호"""
    last = 1
    for row in ws.iter_rows(min_row=1, max_row=9999):
        if any(c.value is not None for c in row):
            last = row[0].row
    return last


def append_to_jd_analysis(ws_jd, parsed, track, link, job_id, collected_date):
    """파싱 결과를 jd_analysis 시트에 행 추가"""
    next_row = get_last_row(ws_jd) + 1
    company   = parsed.get("회사명", "")
    job_title = parsed.get("직무명", "")
    업무_text  = "\n".join(f"• {item}" for item in parsed.get("업무", []))
    연차      = parsed.get("연차", "")
    근무형태  = parsed.get("근무형태", "")
    근무지    = parsed.get("근무지", "")

    items = parsed.get("자격요건", [])
    if not items:
        return 0

    for i, item in enumerate(items):
        r = next_row + i
        ws_jd.cell(row=r, column=COL_JD["job_id"],    value=job_id)
        ws_jd.cell(row=r, column=COL_JD["company"],   value=company)
        ws_jd.cell(row=r, column=COL_JD["track"],     value=track)
        ws_jd.cell(row=r, column=COL_JD["job_title"], value=job_title)
        if i == 0:
            ws_jd.cell(row=r, column=COL_JD["업무"],  value=업무_text)
        ws_jd.cell(row=r, column=COL_JD["구분"],      value=item["구분"])
        ws_jd.cell(row=r, column=COL_JD["내용"],      value=item["내용"])
        ws_jd.cell(row=r, column=COL_JD["연차"],      value=연차)
        ws_jd.cell(row=r, column=COL_JD["link"],      value=link)
        ws_jd.cell(row=r, column=COL_JD["근무형태"],  value=근무형태)
        ws_jd.cell(row=r, column=COL_JD["근무지"],    value=근무지)
        ws_jd.cell(row=r, column=COL_JD["collected"], value=collected_date)

    return len(items)


def main():
    if not API_KEY:
        print("❌ ANTHROPIC_API_KEY 환경변수가 설정되지 않았습니다.")
        print("   export ANTHROPIC_API_KEY='sk-ant-...' 후 다시 실행하세요.")
        return

    print(f"📂 파일 열기: {EXCEL_PATH}")
    wb     = openpyxl.load_workbook(EXCEL_PATH, keep_vba=True)
    ws_raw = wb['raw_data']
    ws_jd  = wb['jd_analysis']
    client = anthropic.Anthropic(api_key=API_KEY)

    # 상태 = '신규'인 행만 수집
    last_row = get_last_row(ws_raw)
    targets  = []
    for r in range(2, last_row + 1):
        status   = ws_raw.cell(row=r, column=COL_RAW["status"]).value
        raw_text = ws_raw.cell(row=r, column=COL_RAW["raw_text"]).value
        if status == "신규" and raw_text:
            targets.append(r)

    if not targets:
        print("✅ 처리할 신규 공고가 없습니다. (raw_data F열 상태 = '신규' 확인)")
        return

    print(f"🔍 신규 공고 {len(targets)}개 발견 → 파싱 시작\n")

    success = 0
    for r in targets:
        job_id    = ws_raw.cell(row=r, column=COL_RAW["job_id"]).value
        track     = ws_raw.cell(row=r, column=COL_RAW["track"]).value
        link      = ws_raw.cell(row=r, column=COL_RAW["link"]).value
        raw_text  = ws_raw.cell(row=r, column=COL_RAW["raw_text"]).value
        collected = ws_raw.cell(row=r, column=COL_RAW["collected"]).value

        print(f"⏳ [{job_id}] 파싱 중...")

        try:
            parsed = parse_jd(client, raw_text)
            count  = append_to_jd_analysis(
                ws_jd, parsed, track, link, job_id, collected
            )
            ws_raw.cell(row=r, column=COL_RAW["status"]).value = "분석완료"
            company   = parsed.get("회사명", "?")
            job_title = parsed.get("직무명", "?")
            print(f"   ✅ {company} — {job_title} | 자격요건 {count}개 추가")
            success += 1

        except json.JSONDecodeError as e:
            print(f"   ❌ JSON 파싱 실패: {e}")
        except Exception as e:
            print(f"   ❌ 오류 발생: {e}")

    wb.save(EXCEL_PATH)
    print(f"\n🎉 완료 — {success}/{len(targets)}개 공고 처리, 파일 저장됨.")


if __name__ == "__main__":
    main()