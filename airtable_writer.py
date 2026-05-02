"""
airtable_writer.py - 파싱 결과 에어테이블 직접 업데이트

사용:
    from airtable_writer import update_parsed_result
    update_parsed_result(name="이윤경", parsed={...})

에어테이블 필드 (접수명단 테이블):
    수입       : 수입금액 (숫자)
    장부유형   : 싱글셀렉트 (간편장부(기준경비율)/복식부기의무자/성실신고확인대상자/...)
    타소득유형 : 싱글셀렉트 (없음/근로소득/연금소득/금융소득/기타소득/근로+기타/...)
"""
import json, urllib.request, urllib.parse
from pathlib import Path

# ── 설정 ──────────────────────────────────────────────
BASE_ID  = "appSvDTDOmYfBeIFs"
TABLE_ID = "tbl2f2h6GfSnLCQpt"

def _get_pat() -> str:
    candidates = [
        Path.home() / "종소세2026/.credentials/airtable_pat.txt",  # Mac Mini
        Path(r"F:\종소세2026\.credentials\airtable_pat.txt"),       # Windows
    ]
    for p in candidates:
        if p.exists():
            return p.read_text().strip()
    raise FileNotFoundError("airtable_pat.txt 없음 — .credentials/ 폴더에 PAT 파일 생성 필요")


# ── 장부유형 → 에어테이블 셀렉트 매핑 ───────────────────
def map_장부유형_select(기장의무: str, 경비율: str = "") -> str | None:
    """파싱결과 기장의무 + 경비율 → 에어테이블 장부유형 셀렉트 값
    매핑 안 되는 경우 None 반환 (기존값 유지, 덮어쓰지 않음)
    """
    기장 = (기장의무 or "").strip()
    경비 = (경비율 or "").strip()

    if "성실신고" in 기장:
        return "성실신고확인대상자"
    if "복식부기" in 기장:
        return "복식부기의무자"
    if "간편" in 기장 and "기준경비율" in 경비:
        return "간편장부(기준경비율)"

    return None  # 나머지는 수동 처리 (안건드림)


# ── 타소득 → 에어테이블 셀렉트 매핑 ─────────────────────
def map_타소득_select(parsed: dict) -> str:
    """파싱결과 딕셔너리 → 에어테이블 싱글셀렉트 값"""
    금융 = parsed.get("이자") == "O" or parsed.get("배당") == "O"
    근로 = parsed.get("근로(단일)") == "O" or parsed.get("근로(복수)") == "O"
    연금 = parsed.get("연금") == "O"
    기타 = parsed.get("기타") == "O"

    mapping = {
        (False, False, False, False): "없음",
        (True,  False, False, False): "근로소득",
        (False, True,  False, False): "연금소득",
        (False, False, True,  False): "금융소득",
        (False, False, False, True ): "기타소득",
        (True,  False, False, True ): "근로+기타",
        (True,  False, True,  False): "근로+금융",
        (True,  True,  False, False): "근로+연금",
        (False, True,  True,  False): "연금+금융",
        (True,  True,  False, True ): "근로+연금+기타",
        (True,  True,  True,  False): "근로+연금+금융",
        (False, True,  False, True ): "연금+기타",
        (False, False, True,  True ): "금융+기타",
        (True,  False, True,  True ): "근로+금융+기타",
        (False, True,  True,  True ): "연금+금융+기타",
        (True,  True,  True,  True ): "근로+연금+금융+기타",
    }
    key = (근로, 연금, 금융, 기타)
    return mapping.get(key, "없음")


# ── 레코드 ID 조회 ────────────────────────────────────
def find_record_id(name: str, pat: str) -> str | None:
    """성명으로 에어테이블 레코드 ID 조회"""
    formula = urllib.parse.quote(f'{{성명}}="{name}"')
    url = f"https://api.airtable.com/v0/{BASE_ID}/{TABLE_ID}?filterByFormula={formula}&maxRecords=1"
    req = urllib.request.Request(url, headers={"Authorization": f"Bearer {pat}"})
    data = json.loads(urllib.request.urlopen(req, timeout=10).read())
    records = data.get("records", [])
    return records[0]["id"] if records else None


# ── 에어테이블 PATCH ──────────────────────────────────
def patch_record(record_id: str, fields: dict, pat: str):
    url = f"https://api.airtable.com/v0/{BASE_ID}/{TABLE_ID}/{record_id}"
    payload = json.dumps({"fields": fields}).encode()
    req = urllib.request.Request(
        url, data=payload, method="PATCH",
        headers={
            "Authorization": f"Bearer {pat}",
            "Content-Type": "application/json",
        }
    )
    resp = urllib.request.urlopen(req, timeout=10)
    return json.loads(resp.read())


# ── 메인 함수 ─────────────────────────────────────────
def update_parsed_result(name: str, parsed: dict) -> bool:
    """파싱 결과를 에어테이블에 직접 업데이트.

    parsed 딕셔너리 예시 (parse_and_sync_신규.py 결과):
        {
            "수입금액총계": 190707661,
            "기장의무": "복식부기의무자",
            "추계시적용경비율": "기준경비율",
            "이자": "", "배당": "", "근로(단일)": "", "근로(복수)": "",
            "연금": "", "기타": "",
        }
    """
    try:
        pat = _get_pat()
        record_id = find_record_id(name, pat)
        if not record_id:
            print(f"  [에어테이블] '{name}' 레코드 없음 — 스킵")
            return False

        타소득_val = map_타소득_select(parsed)

        fields = {}
        if parsed.get("수입금액총계"):
            fields["수입"] = int(parsed["수입금액총계"])
        장부_val = map_장부유형_select(parsed.get("기장의무", ""), parsed.get("추계시적용경비율", ""))
        if 장부_val:
            fields["장부유형"] = 장부_val  # None이면 기존값 유지 (덮어쓰지 않음)
        fields["타소득유형"] = 타소득_val

        patch_record(record_id, fields, pat)
        print(f"  [에어테이블] {name} 업데이트 완료: 수입={fields.get('수입'):,} / {장부_val} / {타소득_val}")
        return True

    except Exception as e:
        print(f"  [에어테이블 오류] {name}: {e}")
        return False


# ── 테스트 ────────────────────────────────────────────
if __name__ == "__main__":
    # 이윤경 테스트 (실제 업데이트 전 record_id 조회만)
    pat = _get_pat()
    rid = find_record_id("이윤경", pat)
    print(f"이윤경 record_id: {rid}")

    # 타소득 매핑 테스트
    cases = [
        ({}, "없음"),
        ({"근로(단일)": "O"}, "근로소득"),
        ({"이자": "O", "배당": "O"}, "금융소득"),
        ({"근로(단일)": "O", "기타": "O"}, "근로+기타"),
        ({"근로(복수)": "O", "연금": "O", "기타": "O"}, "근로+연금+기타"),
    ]
    print("\n[타소득 매핑 테스트]")
    for parsed, expected in cases:
        result = map_타소득_select(parsed)
        ok = "✅" if result == expected else "❌"
        print(f"  {ok} {parsed} → {result} (예상: {expected})")
