"""
step7_consent.py - 수임동의 batch 자동화 (등록 + 폴링 통합)

[2트랙 분기 룰]
  1트랙 (동의만):    등록 시도 → 바로 완료 → 고객에게 "동의해주세요" 안내
  2트랙 (해임후동의): 등록 시도 → 타세무대리인 차단 → "기존 해임 후 동의" 안내
  우리수임:          이미 우리 수임 → skip
  입력오류:          주민번호 등 오류 → 수동 확인

[해지일자 ULTRA CRITICAL]
  해지 후 동의를 같은 날 못 함 (홈택스 시스템 제약)
  → 고객에게 해지일자 = 어제, 동의일자 = 오늘로 안내

원칙:
- 사고는 엉뚱한 데서 발생 → 입력 검증 + sanity check
- 1명당 5초 sleep (국세청 의심 회피)
- dry-run 모드 디폴트 (등록 클릭 안 함)
- 자신없는 alert = "확인필요" 태그
- 카카오 발송은 별도 (이 스크립트는 메시지 생성만)

사전 조건:
- 엣지 디버그 모드 + 세무사 로그인
- 구글시트 '접수명단' 시트에 명단 입력

실행:
  python step7_consent.py            # dry-run (입력만, 등록 X)
  python step7_consent.py --real     # 실등록
  python step7_consent.py --poll     # 폴링만 (등록 안 함)
  python step7_consent.py --real --name 홍길동   # 1명만 실등록
"""
from playwright.sync_api import sync_playwright
from pathlib import Path
import time
import argparse
import sys
import re
import warnings

sys.path.insert(0, r"F:\종소세2026")
from gsheet_writer import (
    get_credentials, get_consent_worksheet,
    load_consent_rows, update_consent_status,
    SPREADSHEET_ID,
)
import gspread

warnings.filterwarnings("ignore")

CONSENT_URL = (
    "https://hometax.go.kr/websquare/websquare.html"
    "?w2xPath=/ui/pp/index_pp.xml"
    "&tmIdx=06&tm2lIdx=0612000000&tm3lIdx=0612010000"
)
STATUS_URL = (
    "https://hometax.go.kr/websquare/websquare.html"
    "?w2xPath=/ui/pp/index_pp.xml"
    "&tmIdx=06&tm2lIdx=0612000000&tm3lIdx=0612040000"
)
SPREADSHEET_ID = "1oh31k00Oa2lZWvu5fnBRVmurdlll1YEG8Fefi5FRfBI"
SLEEP_BETWEEN = 5  # 1명당 sleep (초)

# 수임상태 값 상수
S_TRACK1 = "1트랙_동의요청"    # 등록 완료 → 동의만 안내
S_TRACK2 = "2트랙_해임후동의"  # 타세무 차단 → 해임 + 동의 안내
S_WE_HAVE = "우리수임완료"     # 이미 우리가 수임 → skip
S_INPUT_ERR = "입력오류"
S_DRY = "DRY확인"
S_ERROR = "에러"
S_CHECK = "확인필요"


# ----- 핸드폰번호 정규화 -----
def normalize_phone(raw):
    """핸드폰번호 → ('010', 'XXXX', 'XXXX') 3분할
    1099956363 → 010-9995-6363
    01099956363 → 010-9995-6363
    """
    s = re.sub(r"[^\d]", "", str(raw))
    if len(s) == 10:
        s = "0" + s  # leading 0 보정
    if len(s) != 11:
        return None
    if not s.startswith("01"):
        return None
    return s[:3], s[3:7], s[7:11]


def normalize_jumin(raw):
    """주민번호 → (앞6, 뒤7)"""
    s = re.sub(r"[^\d]", "", str(raw))
    if len(s) != 13:
        return None
    return s[:6], s[6:]


# ----- 2트랙 분기 -----
def classify_alert(alert_msg: str) -> str:
    """
    alert 텍스트 → 수임상태 분류

    1트랙 신호: "등록이 완료" (등록 성공)
    2트랙 신호: "정보제공범위" + "등록되어" (타세무대리인 차단)
               OR "타세무" 포함
    우리수임:   "해임 후 동의" (우리가 이미 수임한 상태에서 재시도)
               OR "이미 수임" 명시적
    입력오류:   "올바르지 않" "입력" 등
    """
    if not alert_msg:
        return S_CHECK + "(alert없음)"

    msg = alert_msg  # 전체 문자열 (여러 alert 누적)

    # 1트랙: 등록 완료
    if "등록이 완료" in msg and "해임" not in msg:
        return S_TRACK1

    # 2트랙: 타세무대리인 수임 차단
    if ("정보제공범위" in msg and "등록되어" in msg):
        return S_TRACK2
    if "타세무대리" in msg:
        return S_TRACK2

    # 우리수임: 우리가 이미 수임한 상태
    if "해임 후 동의" in msg:
        return S_WE_HAVE
    if "이미 수임" in msg:
        return S_WE_HAVE

    # 입력 오류
    if "올바르지 않" in msg or "입력" in msg and "오류" in msg:
        return S_INPUT_ERR

    # 판단 불가
    return S_CHECK


def generate_track1_message(customer: dict) -> str:
    """1트랙: 등록 완료 → 고객에게 동의 요청 안내"""
    name = customer.get("성명", "")
    return f"""[세무회계창연] 수임동의 요청 안내

{name}님 안녕하세요.
2024 귀속 종합소득세 신고를 위해 수임 등록을 완료했습니다.

아래 방법으로 동의해 주시면 신고가 진행됩니다.

▶ 홈택스 앱 동의 방법
  1. 홈택스 앱 로그인 (또는 PC 홈택스)
  2. [세금신고] → [수임동의] → [동의하기]
  3. 세무회계창연 확인 후 동의

※ 동의 완료 후 별도 안내드리겠습니다.
문의: 세무회계창연"""


def generate_track2_message(customer: dict) -> str:
    """2트랙: 타세무대리인 차단 → 해임 후 동의 안내"""
    from datetime import datetime, timedelta
    name = customer.get("성명", "")
    yesterday = (datetime.now() - timedelta(days=1)).strftime("%Y년 %m월 %d일")
    today = datetime.now().strftime("%Y년 %m월 %d일")
    return f"""[세무회계창연] 수임변경 안내

{name}님 안녕하세요.
기존 세무대리인 수임동의가 등록되어 있어 변경이 필요합니다.

▶ 처리 방법 (홈택스 앱)
  1. 홈택스 앱 로그인
  2. [세금신고] → [수임동의] → 기존 세무대리인 해지
     ⚠️ 해지일자를 반드시 [{yesterday}] 로 입력해 주세요
        (오늘 날짜로 해지하면 내일까지 동의 불가)
  3. 세무회계창연 동의
     (수임일자: {today})

완료 후 신고 진행하겠습니다.
문의: 세무회계창연"""


# ----- 시트 로딩 -----
def load_customers():
    """미동의명단 시트 로딩 (step4/5 에러 = 미동의 대상)"""
    ws = get_consent_worksheet()
    rows = load_consent_rows()
    return ws, rows


# ----- 등록 화면 동작 -----
def fill_consent_form(page, customer):
    """수임납세자 등록 화면에 입력 (등록 클릭은 별도)"""
    print(f"  [입력] {customer['성명']} ({customer['주민번호']})", flush=True)

    # 페이지 풀 렌더링 대기 (websquare는 느림)
    print(f"    페이지 로드 대기 5초", flush=True)
    time.sleep(5)

    # 진단: 보이는 input·라디오·텍스트 dump
    diag = page.evaluate("""
        () => {
            const result = {radios: [], inputs: [], hasNonBiz: false};
            for (const r of document.querySelectorAll('input[type=radio]')) {
                if (r.offsetParent === null) continue;
                let label = '';
                const id = r.id;
                if (id) {
                    const lab = document.querySelector(`label[for=${id}]`);
                    if (lab) label = lab.innerText.trim();
                }
                if (!label) {
                    const next = r.nextElementSibling;
                    if (next) label = (next.innerText || '').trim().slice(0,20);
                }
                result.radios.push({id: r.id, name: r.name, label, value: r.value});
            }
            const txt = document.body.innerText || '';
            result.hasNonBiz = txt.includes('비사업자');
            return result;
        }
    """)
    print(f"    [진단] 비사업자 텍스트 페이지에 보임: {diag['hasNonBiz']}")
    print(f"    [진단] 보이는 라디오 {len(diag['radios'])}개:")
    for r in diag['radios'][:10]:
        print(f"        {r}")

    # 1. 비사업자 라디오 클릭 - 라벨/value/id 다양하게 시도
    nonbiz = page.evaluate("""
        () => {
            // 1차: label[for] → 비사업자 텍스트
            for (const lab of document.querySelectorAll('label')) {
                if ((lab.innerText || '').trim() === '비사업자' && lab.offsetParent !== null) {
                    lab.click();
                    return 'label';
                }
            }
            // 2차: 라디오 next sibling이 비사업자
            for (const r of document.querySelectorAll('input[type=radio]')) {
                if (r.offsetParent === null) continue;
                const next = r.nextElementSibling;
                if (next && (next.innerText || '').trim() === '비사업자') {
                    r.click();
                    return 'radio_next';
                }
            }
            // 3차: 라디오 value 비사업자
            for (const r of document.querySelectorAll('input[type=radio]')) {
                if (r.offsetParent === null) continue;
                if ((r.value || '').includes('비사업자') || (r.value || '').toLowerCase().includes('non')) {
                    r.click();
                    return 'value';
                }
            }
            // 4차: span에 비사업자
            for (const s of document.querySelectorAll('span, div, td')) {
                if ((s.innerText || '').trim() === '비사업자' && s.offsetParent !== null) {
                    s.click();
                    return 'span';
                }
            }
            return false;
        }
    """)
    print(f"    비사업자 클릭 결과: {nonbiz}", flush=True)
    time.sleep(1)

    # 2. 주민번호 입력
    jumin = normalize_jumin(customer["주민번호"])
    if not jumin:
        raise ValueError(f"주민번호 형식 이상: {customer['주민번호']}")

    # 주민등록번호 라벨 옆 input (text/password 만)
    jumin_inputs = [el for el in page.locator(
        "xpath=//th[contains(.,'주민등록번호')]/following-sibling::td//input[@type='text' or @type='password']"
    ).all() if el.is_visible()]
    print(f"    주민번호 input {len(jumin_inputs)}개", flush=True)
    # 각 input의 maxlength·type 진단
    for i, el in enumerate(jumin_inputs):
        ml = el.get_attribute("maxlength") or "?"
        tp = el.get_attribute("type") or "?"
        print(f"      [{i}] type={tp} maxlength={ml}", flush=True)

    if not jumin_inputs:
        raise RuntimeError("주민번호 input 못 찾음")

    # 사람처럼 첫 input에 click + keyboard.type으로 13자리 입력
    # (홈택스가 사람 입력 시 자동으로 input 분할 처리)
    first_input = jumin_inputs[0]
    first_input.click()
    page.keyboard.press("Control+a")  # 기존 값 선택
    page.keyboard.press("Delete")
    page.keyboard.type(jumin[0] + jumin[1], delay=50)
    print(f"      → keyboard.type 으로 13자리 입력", flush=True)

    # 입력값 검증
    for i, el in enumerate(jumin_inputs):
        try:
            v = el.input_value()
            print(f"      [{i}] 입력값='{v}' (길이 {len(v)})", flush=True)
        except Exception:
            pass

    # 3. 성명
    name_input = page.locator(
        "xpath=//th[contains(.,'성명')]/following-sibling::td//input"
    ).first
    name_input.fill(customer["성명"])

    # 4. 전화번호 (3분할) - type=text 만
    phone = normalize_phone(customer["핸드폰번호"])
    if not phone:
        raise ValueError(f"핸드폰번호 형식 이상: {customer['핸드폰번호']}")

    tel_inputs = [el for el in page.locator(
        "xpath=//th[contains(.,'전화번호') and not(contains(.,'휴대'))]/following-sibling::td//input[@type='text']"
    ).all() if el.is_visible()]
    print(f"    전화번호 input {len(tel_inputs)}개", flush=True)
    if len(tel_inputs) >= 3:
        # 첫 input은 select일 수도 (010 드롭다운)
        try:
            # select tag 시도
            tel_select = page.locator(
                "xpath=//th[contains(.,'전화번호') and not(contains(.,'휴대'))]/following-sibling::td//select"
            ).all()
            if tel_select:
                page.evaluate(f"""
                    (sel) => {{ sel.value = '{phone[0]}'; sel.dispatchEvent(new Event('change')); }}
                """, tel_select[0].element_handle())
                # 두 번째·세 번째 input
                if len(tel_inputs) >= 2:
                    tel_inputs[-2].fill(phone[1])
                    tel_inputs[-1].fill(phone[2])
            else:
                tel_inputs[0].fill(phone[0])
                tel_inputs[1].fill(phone[1])
                tel_inputs[2].fill(phone[2])
        except Exception as e:
            print(f"    전화번호 입력 예외: {e}", flush=True)

    # 5. 휴대전화번호 - type=text 만 (라디오 제외)
    mobile_inputs = [el for el in page.locator(
        "xpath=//th[contains(.,'휴대전화번호')]/following-sibling::td//input[@type='text']"
    ).all() if el.is_visible()]
    print(f"    휴대전화 input {len(mobile_inputs)}개", flush=True)
    if len(mobile_inputs) >= 2:
        # 마지막 2개 input에 4자리/4자리 채움
        mobile_inputs[-2].fill(phone[1])
        mobile_inputs[-1].fill(phone[2])
        # select가 있으면 010으로 설정
        try:
            mobile_select = [s for s in page.locator(
                "xpath=//th[contains(.,'휴대전화번호')]/following-sibling::td//select"
            ).all() if s.is_visible()]
            if mobile_select:
                mobile_select[0].select_option(phone[0])
        except Exception as e:
            print(f"    휴대전화 select 예외 (무시): {e}", flush=True)

    # 6. 수임일자 = 오늘
    from datetime import datetime
    today = datetime.now().strftime("%Y-%m-%d")
    try:
        date_input = page.locator(
            "xpath=//th[contains(.,'수임일자')]/following-sibling::td//input"
        ).first
        date_input.fill(today)
        # 또는 placeholder yyyy-mm-dd 직접
        filled_date = date_input.input_value()
        print(f"    수임일자 입력: {filled_date}", flush=True)
    except Exception as e:
        print(f"    수임일자 입력 예외: {e}", flush=True)

    print(f"  [입력 완료]", flush=True)


def click_register_and_capture_alert(page):
    """등록하기 클릭 + 모든 alert 누적 캡처
    - 1차 alert: 'XXX 등록 하시겠습니까?' (confirm) → accept
    - 2차 alert: '등록이 완료되었습니다...' 또는 거부 메시지 → 캡처
    """
    alert_text = []
    def on_dialog(d):
        msg = d.message
        alert_text.append(msg)
        print(f"      [dialog 받음] {msg[:80]}", flush=True)
        d.accept()
    page.on("dialog", on_dialog)

    page.evaluate("""
        () => {
            for (const el of document.querySelectorAll('input[type=button], button, a')) {
                const v = (el.value || el.innerText || '').trim();
                if (v === '등록하기' && el.offsetParent !== null) {
                    el.click();
                    return true;
                }
            }
            return false;
        }
    """)

    # 두 번째 alert (결과) 기다리기 - 최대 10초 polling
    for _ in range(20):
        time.sleep(0.5)
        if len(alert_text) >= 2:
            break

    return "\n".join(alert_text) if alert_text else ""


# ----- 메인 -----
def parse_args():
    p = argparse.ArgumentParser()
    p.add_argument("--real", action="store_true", help="실제 등록 (등록하기 클릭)")
    p.add_argument("--poll", action="store_true", help="폴링만 (등록 안 함)")
    p.add_argument("--name", help="특정 고객 1명만 (성명)")
    return p.parse_args()


def main():
    args = parse_args()
    is_dry = not args.real
    do_register = not args.poll

    print(f"\n[모드] 등록={'실등록' if not is_dry else 'DRY-RUN'} 폴링={'on' if not args.real else 'off'}")
    print(f"[안전] 1명당 {SLEEP_BETWEEN}초 sleep, alert 자동 캡처\n")

    ws, customers = load_customers()
    print(f"[접수명단] 총 {len(customers)}명")

    # 시트 row index 추적 (헤더 제외 시 row 2부터)
    customers_with_idx = [(i, c) for i, c in enumerate(customers, start=2)]

    if args.name:
        customers_with_idx = [(i, c) for i, c in customers_with_idx if c.get("성명") == args.name]
        print(f"[필터] '{args.name}' 만 처리 → {len(customers_with_idx)}명")
        if not customers_with_idx:
            print(f"[에러] '{args.name}' 명단에 없음")
            return

    if do_register:
        with sync_playwright() as p:
            browser = p.chromium.connect_over_cdp("http://localhost:9222")
            ctx = browser.contexts[0]
            page = ctx.pages[0]
            page.bring_to_front()

            for n, (idx, c) in enumerate(customers_with_idx, start=1):
                print(f"\n[{n}/{len(customers_with_idx)}] {c.get('성명')} (시트행 {idx})")
                # 처리 대상: 조회 에러 = 미동의 상태만
                # (조회 성공한 사람은 이미 수임완료 → 여기 올 일 없음)
                if not args.name and c.get("수임상태") not in ("", "미동의", "미신청", None):
                    print(f"    skip (이미 {c.get('수임상태')})")
                    continue

                try:
                    page.goto(CONSENT_URL, wait_until="domcontentloaded")
                    time.sleep(3)
                    print(f"    [진단] goto 후 URL: {page.url}")
                    # 페이지 제목·내용 확인
                    page_check = page.evaluate("""
                        () => ({
                            title: document.title,
                            hasRegister: (document.body.innerText || '').includes('수임납세자 등록'),
                            hasLogin: (document.body.innerText || '').includes('아이디 로그인') || (document.body.innerText || '').includes('공동·금융 인증'),
                        })
                    """)
                    print(f"    [진단] {page_check}")
                    if page_check.get("hasLogin"):
                        raise RuntimeError("로그인 페이지로 redirect됨 - 세션 만료?")
                    if not page_check.get("hasRegister"):
                        print(f"    [진단] '수임납세자 등록' 텍스트 안 보임 - 메뉴 클릭으로 fallback")
                        # 메뉴 클릭: 세무대리/납세관리 → 기장·수임 납세자관리 → 기장대리 수임납세자 등록
                        page.locator("text=세무대리/납세관리").first.hover()
                        time.sleep(2)
                        page.evaluate("""
                            () => {
                                for (const el of document.querySelectorAll('a, span, div, li')) {
                                    if ((el.innerText || '').trim() === '기장대리 수임납세자 등록' && el.offsetParent !== null) {
                                        el.click();
                                        return true;
                                    }
                                }
                                return false;
                            }
                        """)
                        time.sleep(3)
                        print(f"    [진단] 메뉴 클릭 후 URL: {page.url}")

                    fill_consent_form(page, c)

                    if is_dry:
                        print(f"  [DRY-RUN] 등록 클릭 안 함 - 5초 화면 확인")
                        time.sleep(5)
                        update_consent_status(ws, idx, S_DRY, "입력만 됨")
                    else:
                        alert_msg = click_register_and_capture_alert(page)
                        print(f"  [Alert 누적]\n{alert_msg}")

                        # ----- 2트랙 분기 -----
                        status = classify_alert(alert_msg)
                        print(f"  [분류] {status}")

                        # 트랙별 카카오 메시지 생성
                        kakao_msg = ""
                        if status == S_TRACK1:
                            kakao_msg = generate_track1_message(c)
                            print(f"  [1트랙] 동의 요청 메시지 생성 완료")
                        elif status == S_TRACK2:
                            kakao_msg = generate_track2_message(c)
                            print(f"  [2트랙] 해임+동의 요청 메시지 생성 완료")
                        elif status == S_WE_HAVE:
                            print(f"  [우리수임] 이미 등록됨 - skip")
                        else:
                            print(f"  [⚠️ 수동확인 필요] status={status}")

                        update_consent_status(ws, idx, status, alert_msg, kakao_msg)

                except Exception as e:
                    print(f"  [에러] {e}")
                    update_consent_status(ws, idx, S_ERROR, str(e)[:300])

                time.sleep(SLEEP_BETWEEN)

    # ----- 완료 후 요약 -----
    print("\n" + "=" * 60)
    print("[등록 단계 완료] 미동의명단 수임상태 요약:")
    latest = load_consent_rows()
    counts = {}
    for c in latest:
        s = c.get("수임상태", "") or ""
        counts[s] = counts.get(s, 0) + 1
    for s, n in sorted(counts.items()):
        print(f"  {s}: {n}명")
    print()

    # 카카오 발송 대기 목록 (1트랙 + 2트랙)
    pending = [c for c in latest if c.get("수임상태") in (S_TRACK1, S_TRACK2)]
    if pending:
        print(f"[카카오 발송 대기] {len(pending)}명:")
        for c in pending:
            print(f"  [{c['수임상태']}] {c['성명']} ({c.get('핸드폰번호','?')})")
    else:
        print("[카카오 발송 대기] 0명")
    print("※ 카카오 발송은 구글시트 '카카오발송문' 열 확인 후 수동 또는 별도 발송 스크립트")


if __name__ == "__main__":
    main()
