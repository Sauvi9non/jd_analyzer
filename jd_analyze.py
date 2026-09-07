"""
공고 분석 에이전트 — 트랙별 역량 패턴 분석
동작: jd_analysis 읽기 → LLM으로 분류/정규화 → jd_insight 시트에 저장
사용법: python3 jd_analyze.py
"""

import os
import json
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
import anthropic
from collections import defaultdict

# ================================================================
# 설정
# ================================================================
EXCEL_PATH = "/Users/maczoo/Documents/Vault/Attached/jd_analysis.xlsm"
API_KEY    = os.environ.get("ANTHROPIC_API_KEY", "")
MODEL      = "claude-sonnet-4-6"

# ================================================================
# 프롬프트
# ================================================================
CLASSIFY_PROMPT = """당신은 채용공고 자격요건 분류 전문가입니다.
아래 자격요건 항목 목록을 읽고 각 항목을 분류하여 JSON으로 반환하세요.
반드시 JSON만 반환하고 마크다운 코드블록은 포함하지 마세요.

분류 기준:
- 기술스택: 특정 도구·언어·프레임워크·플랫폼 (SQL, Python, dbt, Spark, AWS, LLM, RAG 등)
- 도메인경험: 특정 도메인 지식이나 업무 경험 (데이터 모델링 경험, 커머스 도메인 이해, A/B테스트 경험 등)
- 소프트스킬: 의사소통, 협업, 문제해결, 주도성 등 일반적 역량

반환 형식:
[
  {
    "원문": "원문 그대로",
    "구분": "필수" 또는 "우대",
    "분류": "기술스택" 또는 "도메인경험" 또는 "소프트스킬",
    "정규화": "정규화된 키워드 또는 카테고리 (기술스택은 도구명, 도메인경험은 간결한 카테고리명, 소프트스킬은 그대로)"
  },
  ...
]

기술스택 정규화 예시:
- "SQL을 능숙하게 다루시는 분" → "SQL"
- "Python 활용 가능자" → "Python"
- "dbt 사용 경험" → "dbt"
- "AWS 또는 GCP 클라우드 경험" → "AWS/GCP"
- "LLM API 사용 경험" → "LLM API"

도메인경험 정규화 예시:
- "5년 이상 데이터 분석 경험" → "데이터 분석 경험"
- "커머스 도메인 데이터 모델링 경험" → "데이터 모델링 경험"
- "A/B 테스트 설계 및 분석 경험" → "A/B 테스트 경험"
- "데이터 파이프라인 구축 경험" → "데이터 파이프라인 경험"

기술스택 정규화 추가 규칙:
- 여러 기술이 '/', '또는', 'or'로 연결된 경우 각각 별도 항목으로 분리해서 반환하세요.
  예: "AWS 또는 GCP 경험" → 두 개의 항목 {"정규화": "AWS"}, {"정규화": "GCP"}
  예: "SQL/Python 활용" → {"정규화": "SQL"}, {"정규화": "Python"}
- 단, 'LLM API', 'A/B 테스트'처럼 하나의 고유 명사인 경우는 분리하지 마세요.

"""

def classify_items(client, items, batch_size=30):
    """자격요건 항목 배치 분류 (batch_size씩 나눠서 처리)"""
    all_results = []
    for start in range(0, len(items), batch_size):
        batch = items[start:start + batch_size]
        items_text = "\n".join(
            f"{i+1}. [{item['구분']}] {item['내용']}"
            for i, item in enumerate(batch)
        )
        response = client.messages.create(
            model=MODEL,
            max_tokens=8000,
            system=CLASSIFY_PROMPT,
            messages=[{"role": "user", "content": f"다음 자격요건을 분류해주세요:\n\n{items_text}"}]
        )
        batch_result = json.loads(response.content[0].text.strip())
        all_results.extend(batch_result)
    return all_results


def load_jd_analysis(ws_jd):
    """jd_analysis 시트에서 트랙별 자격요건 수집"""
    last_row = 1
    for row in ws_jd.iter_rows(min_row=1, max_row=9999):
        if any(c.value is not None for c in row):
            last_row = row[0].row

    track_data = defaultdict(lambda: {"job_ids": set(), "items": []})

    for r in range(2, last_row + 1):
        job_id  = ws_jd.cell(row=r, column=1).value
        track   = ws_jd.cell(row=r, column=3).value
        구분    = ws_jd.cell(row=r, column=7).value
        내용    = ws_jd.cell(row=r, column=8).value
        if track and 구분 and 내용:
            track_data[track]["job_ids"].add(job_id)
            track_data[track]["items"].append({"구분": 구분, "내용": 내용, "job_id": job_id})

    return track_data


def aggregate(classified, job_ids_per_item, total_companies):
    """분류 결과 집계 — 기술스택/도메인경험별 언급 공고 수 및 비율"""
    from collections import defaultdict
    agg = defaultdict(lambda: {"필수": set(), "우대": set()})

    for item, job_id in zip(classified, job_ids_per_item):
        if item["분류"] == "소프트스킬":
            continue
        key = (item["분류"], item["정규화"])
        agg[key][item["구분"]].add(job_id)

    results = []
    for (분류, 정규화), counts in agg.items():
        필수_n = len(counts["필수"])
        우대_n = len(counts["우대"])
        total  = len(counts["필수"] | counts["우대"])
        results.append({
            "분류": 분류,
            "키워드": 정규화,
            "필수_공고수": 필수_n,
            "우대_공고수": 우대_n,
            "총_공고수": total,
            "비율": round(total / total_companies * 100),
        })

    # 총 공고수 내림차순 정렬, 같으면 필수 우선
    results.sort(key=lambda x: (-x["총_공고수"], -x["필수_공고수"]))
    return results


def write_insight_sheet(wb, analysis_by_track):
    """jd_insight 시트 생성 및 결과 작성"""
    if "jd_insight" in wb.sheetnames:
        del wb["jd_insight"]
    ws = wb.create_sheet("jd_insight")

    # 헤더 스타일
    HEADER_FILL = PatternFill("solid", fgColor="FF2D2D2D")
    HEADER_FONT = Font(name="맑은 고딕", bold=True, color="FFFFFFFF", size=11)
    TRACK_FILLS = {"A": "FFE8F0FE", "B": "FFFCE8E6", "C": "FFE6F4EA"}

    headers = ["트랙", "분류", "키워드/역량", "필수 공고수", "우대 공고수", "총 공고수", "비율(%)"]
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")

    col_widths = [8, 14, 32, 12, 12, 12, 10]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[chr(64+i)].width = w
    ws.row_dimensions[1].height = 22
    ws.freeze_panes = "A2"

    row = 2
    for track in ["A", "B", "C"]:
        if track not in analysis_by_track:
            continue
        results       = analysis_by_track[track]["results"]
        total_co      = analysis_by_track[track]["total_companies"]
        soft_skills   = analysis_by_track[track]["soft_skills"]
        fill = PatternFill("solid", fgColor=TRACK_FILLS[track])

        for item in results:
            ws.cell(row=row, column=1, value=f"트랙 {track}").fill = fill
            ws.cell(row=row, column=2, value=item["분류"])
            ws.cell(row=row, column=3, value=item["키워드"])
            ws.cell(row=row, column=4, value=item["필수_공고수"])
            ws.cell(row=row, column=5, value=item["우대_공고수"])
            ws.cell(row=row, column=6, value=item["총_공고수"])
            ws.cell(row=row, column=7, value=f'{item["비율"]}%')
            ws.cell(row=row, column=1).alignment = Alignment(horizontal="center")
            ws.cell(row=row, column=7).alignment = Alignment(horizontal="center")
            row += 1

        # 소프트스킬 요약 행
        if soft_skills:
            soft_text = " / ".join(sorted(set(soft_skills))[:8])
            ws.cell(row=row, column=1, value=f"트랙 {track}").fill = fill
            ws.cell(row=row, column=2, value="소프트스킬")
            ws.cell(row=row, column=3, value=f"[공통] {soft_text}")
            ws.cell(row=row, column=4, value="-")
            ws.cell(row=row, column=5, value="-")
            ws.cell(row=row, column=6, value="-")
            ws.cell(row=row, column=7, value="-")
            ws.cell(row=row, column=1).alignment = Alignment(horizontal="center")
            for c in range(1, 8):
                ws.cell(row=row, column=c).font = Font(name="맑은 고딕", italic=True, color="FF888888")
            row += 2  # 트랙 간 빈 행

    print(f"   jd_insight 시트 작성 완료 ({row-1}행)")


def main():
    if not API_KEY:
        print("❌ ANTHROPIC_API_KEY 환경변수가 설정되지 않았습니다.")
        return

    print(f"📂 파일 열기: {EXCEL_PATH}")
    wb    = openpyxl.load_workbook(EXCEL_PATH, keep_vba=True)
    ws_jd = wb["jd_analysis"]

    print("📊 jd_analysis 데이터 로딩...")
    track_data = load_jd_analysis(ws_jd)

    client = anthropic.Anthropic(api_key=API_KEY)
    analysis_by_track = {}

    for track in ["A", "B", "C"]:
        if track not in track_data:
            print(f"⚠ 트랙 {track} 데이터 없음, 건너뜀")
            continue

        items         = track_data[track]["items"]
        total_cos     = len(track_data[track]["job_ids"])
        print(f"\n⏳ 트랙 {track} 분류 중... ({len(items)}개 항목, {total_cos}개 공고)")

        try:
            classified = classify_items(client, items)
            job_ids    = [item["job_id"] for item in items]
            results    = aggregate(classified, job_ids, total_cos)
            soft_skills = [
                item["정규화"] for item in classified
                if item["분류"] == "소프트스킬"
            ]
            analysis_by_track[track] = {
                "results": results,
                "total_companies": total_cos,
                "soft_skills": soft_skills,
            }
            tech = sum(1 for r in results if r["분류"] == "기술스택")
            domain = sum(1 for r in results if r["분류"] == "도메인경험")
            print(f"   ✅ 기술스택 {tech}개, 도메인경험 {domain}개, 소프트스킬 {len(soft_skills)}개")

        except json.JSONDecodeError as e:
            print(f"   ❌ JSON 파싱 실패: {e}")
        except Exception as e:
            print(f"   ❌ 오류 발생: {e}")

    if analysis_by_track:
        print("\n📝 jd_insight 시트 작성 중...")
        write_insight_sheet(wb, analysis_by_track)
        wb.save(EXCEL_PATH)
        print(f"\n🎉 완료 — 파일 저장됨: {EXCEL_PATH}")
    else:
        print("\n❌ 분석 결과가 없어 저장하지 않았습니다.")


if __name__ == "__main__":
    main()