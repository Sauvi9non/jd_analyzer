"""
JD Analyzer 파싱 스크립트 v1
- raw_data에서 상태='신규'인 행 읽기
- LLM이 DA/DE/AE 직군 자동 추천 → B열(직군_추천)에 저장
- C열(직군_확정)이 채워진 행만 jd_analysis로 파싱
사용법:
  # 1단계: 직군 추천만 실행
  python3 jd_parse.py --classify

  # 2단계: 직군 확정 후 파싱
  python3 jd_parse.py --parse

  # 둘 다 한번에
  python3 jd_parse.py
"""

import os
import sys
import json
import openpyxl
import anthropic

# ================================================================
# 설정
# ================================================================
EXCEL_PATH = os.path.expanduser("/Users/maczoo/Desktop/jd_analyzer/jd_analyzer/jd_analyzer.xlsx")
API_KEY    = os.environ.get("ANTHROPIC_API_KEY", "")
MODEL      = "claude-sonnet-4-6"

ROLES = {
    "DA": "데이터 분석가 — 지표 분석, A/B 테스트, 비즈니스 인사이트 도출 중심",
    "DE": "데이터 엔지니어 — 파이프라인 구축, ETL, 데이터 인프라 중심",
    "AE": "애널리틱스 엔지니어 — dbt, 데이터 모델링, 분석 인프라와 분석 사이",
    "TW": "테크니컬 라이터 — 기술 문서 작성, UX 라이팅, 개발자 가이드 중심",
}

# ================================================================
# raw_data 컬럼 위치
# ================================================================
COL_RAW = {
    "job_id":     1,  # A
    "role_sug":   2,  # B: 직군_추천
    "role_conf":  3,  # C: 직군_확정
    "link":       4,  # D
    "raw_text":   5,  # E
    "collected":  6,  # F
    "status":     7,  # G
}

# jd_analysis 컬럼 위치
COL_JD = {
    "job_id":    1,   # A
    "company":   2,   # B
    "job_role":  3,   # C
    "job_title": 4,   # D
    "tasks":     5,   # E
    "category":  6,   # F
    "content":   7,   # G
    "seniority": 8,   # H
    "link":      9,   # I
    "work_type": 10,  # J
    "location":  11,  # K
    "collected": 12,  # L
}

# ================================================================
# 프롬프트
# ================================================================
CLASSIFY_PROMPT = f"""당신은 채용공고 직군 분류 전문가입니다.
채용공고 원문을 읽고 직군을 분류하세요.

아래 후보 직군 중 맞는 것이 있으면 해당 코드를 사용하고,
해당하지 않으면 직군명을 간결하게 직접 작성하세요 (예: "PM", "ML Engineer", "DevOps").

후보 직군:
- DA: {ROLES['DA']}
- DE: {ROLES['DE']}
- AE: {ROLES['AE']}
- TW: {ROLES['TW']}

반드시 JSON만 반환하고 마크다운 코드블록은 포함하지 마세요.
반환 형식:
{{
  "직군": "DA" 또는 "DE" 또는 "AE" 또는 "TW" 또는 직접 작성한 직군명,
  "이유": "한 문장으로 분류 근거"
}}"""

PARSE_PROMPT = """당신은 채용공고 파싱 전문가입니다.
채용공고 원문을 읽고 아래 JSON 형식으로 정확히 파싱해서 반환하세요.
반드시 JSON만 반환하고 마크다운 코드블록은 포함하지 마세요.

반환 형식:
{
  "회사명": "회사명 (원문에서 추출, 없으면 빈 문자열)",
  "직무명": "직무명/포지션명",
  "업무": ["업무 항목1", "업무 항목2"],
  "자격요건": [
    {"구분": "필수", "내용": "항목 내용"},
    {"구분": "우대", "내용": "항목 내용"}
  ],
  "연차": "연차 정보 (없으면 빈 문자열)",
  "근무형태": "정규직/계약직/인턴 등 (없으면 빈 문자열)",
  "근무지": "근무 위치 (없으면 빈 문자열)"
}

규칙:
- 업무와 자격요건(필수/우대)을 명확히 구분하세요.
- 자격요건이 필수/우대 구분 없으면 문맥으로 판단해 분류하세요.
- 각 항목은 하나의 내용만 담아주세요."""


def get_last_row(ws):
    last = 1
    for row in ws.iter_rows(min_row=1, max_row=9999):
        if any(c.value is not None for c in row):
            last = row[0].row
    return last


def strip_codeblock(text: str) -> str:
    """마크다운 코드블록 제거"""
    text = text.strip()
    if text.startswith("```"):
        lines = text.split("\n")
        # 첫 줄(```json 등)과 마지막 줄(```) 제거
        lines = lines[1:] if lines[0].startswith("```") else lines
        lines = lines[:-1] if lines and lines[-1].strip() == "```" else lines
        text = "\n".join(lines).strip()
    return text


def classify_role(client, raw_text):
    resp = client.messages.create(
        model=MODEL,
        max_tokens=200,
        system=CLASSIFY_PROMPT,
        messages=[{"role": "user", "content": f"다음 공고를 분류해주세요:\n\n{raw_text[:3000]}"}]
    )
    raw_response = strip_codeblock(resp.content[0].text)
    if not raw_response:
        raise ValueError("API 응답이 비어있습니다.")
    return json.loads(raw_response)


def parse_jd(client, raw_text):
    resp = client.messages.create(
        model=MODEL,
        max_tokens=3000,
        system=PARSE_PROMPT,
        messages=[{"role": "user", "content": f"다음 채용공고를 파싱해주세요:\n\n{raw_text}"}]
    )
    return json.loads(strip_codeblock(resp.content[0].text))


def run_classify(ws_raw, client):
    """1단계: 직군_추천 컬럼 채우기 (상태='신규'이고 직군_추천이 비어있는 행)"""
    last_row = get_last_row(ws_raw)
    targets = []
    for r in range(2, last_row + 1):
        status   = ws_raw.cell(row=r, column=COL_RAW["status"]).value
        raw_text = ws_raw.cell(row=r, column=COL_RAW["raw_text"]).value
        sug      = ws_raw.cell(row=r, column=COL_RAW["role_sug"]).value
        if status == "신규" and raw_text and not sug:
            targets.append(r)

    if not targets:
        print("✅ 직군 추천할 신규 공고가 없습니다.")
        return 0

    print(f"🔍 직군 추천 대상: {len(targets)}개\n")
    count = 0
    for r in targets:
        job_id   = ws_raw.cell(row=r, column=COL_RAW["job_id"]).value
        raw_text = ws_raw.cell(row=r, column=COL_RAW["raw_text"]).value
        print(f"⏳ [{job_id}] 직군 분류 중...")
        try:
            result = classify_role(client, raw_text)
            role   = result.get("직군", "DA")
            reason = result.get("이유", "")
            ws_raw.cell(row=r, column=COL_RAW["role_sug"]).value = role
            print(f"   ✅ {role} — {reason}")
            count += 1
        except Exception as e:
            print(f"   ❌ 오류: {e}")
    return count


def run_parse(ws_raw, ws_jd, client):
    """2단계: 직군_확정된 행을 파싱해서 jd_analysis에 추가"""
    last_row = get_last_row(ws_raw)
    targets = []
    for r in range(2, last_row + 1):
        status    = ws_raw.cell(row=r, column=COL_RAW["status"]).value
        raw_text  = ws_raw.cell(row=r, column=COL_RAW["raw_text"]).value
        role_conf = ws_raw.cell(row=r, column=COL_RAW["role_conf"]).value
        if status == "신규" and raw_text and role_conf:
            targets.append(r)

    if not targets:
        print("✅ 파싱할 공고가 없습니다. (직군_확정 컬럼이 채워진 신규 행 확인)")
        return 0

    print(f"📝 파싱 대상: {len(targets)}개\n")
    count = 0
    for r in targets:
        job_id    = ws_raw.cell(row=r, column=COL_RAW["job_id"]).value
        role_conf = ws_raw.cell(row=r, column=COL_RAW["role_conf"]).value
        raw_text  = ws_raw.cell(row=r, column=COL_RAW["raw_text"]).value
        link      = ws_raw.cell(row=r, column=COL_RAW["link"]).value
        collected = ws_raw.cell(row=r, column=COL_RAW["collected"]).value

        print(f"⏳ [{job_id}] 파싱 중...")
        try:
            parsed    = parse_jd(client, raw_text)
            next_row  = get_last_row(ws_jd) + 1
            tasks_text = "\n".join(f"• {t}" for t in parsed.get("업무", []))
            items     = parsed.get("자격요건", [])

            for i, item in enumerate(items):
                r2 = next_row + i
                ws_jd.cell(row=r2, column=COL_JD["job_id"],    value=job_id)
                ws_jd.cell(row=r2, column=COL_JD["company"],   value=parsed.get("회사명", ""))
                ws_jd.cell(row=r2, column=COL_JD["job_role"],  value=role_conf)
                ws_jd.cell(row=r2, column=COL_JD["job_title"], value=parsed.get("직무명", ""))
                ws_jd.cell(row=r2, column=COL_JD["tasks"],     value=tasks_text if i == 0 else "")
                ws_jd.cell(row=r2, column=COL_JD["category"],  value=item["구분"])
                ws_jd.cell(row=r2, column=COL_JD["content"],   value=item["내용"])
                ws_jd.cell(row=r2, column=COL_JD["seniority"], value=parsed.get("연차", ""))
                ws_jd.cell(row=r2, column=COL_JD["link"],      value=link)
                ws_jd.cell(row=r2, column=COL_JD["work_type"], value=parsed.get("근무형태", ""))
                ws_jd.cell(row=r2, column=COL_JD["location"],  value=parsed.get("근무지", ""))
                ws_jd.cell(row=r2, column=COL_JD["collected"], value=collected)

            ws_raw.cell(row=r, column=COL_RAW["status"]).value = "분석완료"
            company = parsed.get("회사명", "?")
            title   = parsed.get("직무명", "?")
            print(f"   ✅ {company} — {title} | 자격요건 {len(items)}개 추가")
            count += 1
        except Exception as e:
            print(f"   ❌ 오류: {e}")
    return count


def main():
    if not API_KEY:
        print("❌ ANTHROPIC_API_KEY 환경변수가 설정되지 않았습니다.")
        print("   export ANTHROPIC_API_KEY='sk-ant-...' 후 다시 실행하세요.")
        return

    mode_classify = "--parse" not in sys.argv
    mode_parse    = "--classify" not in sys.argv

    print(f"📂 파일 열기: {EXCEL_PATH}")
    wb     = openpyxl.load_workbook(EXCEL_PATH)#, keep_vba=True)
    ws_raw = wb['raw_data']
    ws_jd  = wb['jd_parsed']
    client = anthropic.Anthropic(api_key=API_KEY)

    total = 0

    if mode_classify:
        print("\n[ 1단계: 직군 추천 ]")
        total += run_classify(ws_raw, client)

    if mode_parse:
        print("\n[ 2단계: 파싱 → jd_analysis ]")
        total += run_parse(ws_raw, ws_jd, client)

    wb.save(EXCEL_PATH)
    print(f"\n🎉 완료 — 총 {total}개 처리, 파일 저장됨.")
    print("\n💡 사용 팁:")
    print("  직군 추천만: python3 jd_parse.py --classify")
    print("  파싱만:      python3 jd_parse.py --parse")
    print("  둘 다:       python3 jd_parse.py")


if __name__ == "__main__":
    main()