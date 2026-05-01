"""
sms_monitor_mac.py - Mac Mini SMS 입금 감지 데몬

실행 위치: Mac Mini ~/종소세2026/
실행 방법: python3 sms_monitor_mac.py

흐름:
  카카오뱅크 입금 SMS → iPhone → Mac Mini Messages.app → chat.db
  → 입금자명+금액 파싱 → 구글시트 접수명단 매칭 → 입금체크 O
  → 솔라피 자료접수 안내 발송

전제조건:
  - iPhone 설정 → 메시지 → 문자 메시지 전달 → Mac Mini 켜짐
  - Mac Mini Terminal Full Disk Access 허용
  - pip3 install gspread google-auth requests
"""

import sqlite3
import shutil
import re
import time
import requests
import json
import os
from datetime import datetime

# ===== 설정 =====
DB_PATH       = os.path.expanduser("~/Library/Messages/chat.db")
TMP_DB        = "/tmp/chat_sms_copy.db"
POLL_SEC      = 60          # 폴링 간격 (초)
SPREADSHEET_ID = "1oh31k00Oa2lZWvu5fnBRVmurdlll1YEG8Fefi5FRfBI"
SHEET_접수명단 = "접수명단"

# 솔라피 설정 (자료접수 안내)
SOLAPI_KEY    = ""          # TODO: 솔라피 API KEY
SOLAPI_SECRET = ""          # TODO: 솔라피 API SECRET
SOLAPI_SENDER = ""          # TODO: 발신번호
SOLAPI_TEMPLATE_자료접수 = ""  # TODO: 알림톡 템플릿 코드

# 카카오뱅크 SMS 파싱 패턴
PATTERNS = [
    re.compile(r"입금\s+([\d,]+)원.*?입금자명[:\s]+([^\n\r]+)", re.DOTALL),
    re.compile(r"입금\s+([\d,]+)원\s+([^\n]+)"),
]


# ===== SMS 파싱 =====
def extract_text(blob):
    """attributedBody blob → 텍스트 추출"""
    if not blob:
        return ""
    text = blob.decode("utf-8", errors="ignore")
    m = re.search(r"NSString\x01\x01\+.(.+?)\x02", text, re.DOTALL)
    return m.group(1).strip() if m else ""


def parse_deposit(msg: str):
    """입금 SMS → (금액, 입금자명) or None"""
    if not msg or "입금" not in msg:
        return None
    for pat in PATTERNS:
        m = pat.search(msg)
        if m:
            try:
                amount = int(m.group(1).replace(",", ""))
                name   = m.group(2).strip()
                return amount, name
            except ValueError:
                continue
    return None


# ===== chat.db 폴링 =====
def get_last_rowid():
    shutil.copy2(DB_PATH, TMP_DB)
    conn = sqlite3.connect(TMP_DB)
    row = conn.execute("SELECT MAX(ROWID) FROM message").fetchone()
    conn.close()
    return row[0] or 0


def get_new_kakao_sms(since_rowid: int):
    """새 카카오뱅크 입금 SMS 조회"""
    shutil.copy2(DB_PATH, TMP_DB)
    conn = sqlite3.connect(TMP_DB)
    rows = conn.execute("""
        SELECT m.ROWID,
               m.text,
               m.attributedBody,
               datetime(m.date/1000000000 + 978307200, 'unixepoch', 'localtime') as dt
        FROM message m
        WHERE m.ROWID > ?
          AND m.is_from_me = 0
        ORDER BY m.date ASC
    """, (since_rowid,)).fetchall()
    conn.close()

    results = []
    for rowid, text, blob, dt in rows:
        msg = text or extract_text(blob)
        if msg and "카카오뱅크" in msg and "입금" in msg:
            results.append((rowid, msg, dt))
    return results


# ===== 구글시트 매칭 =====
CRED_DIR      = os.path.expanduser("~/종소세2026/.credentials")
CLIENT_SECRET = os.path.join(CRED_DIR, "client_secret.json")
TOKEN_FILE    = os.path.join(CRED_DIR, "token.pickle")
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

def get_credentials():
    import pickle
    import gspread
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials as OAuthCreds

    creds = None
    if os.path.exists(TOKEN_FILE):
        with open(TOKEN_FILE, "rb") as f:
            creds = pickle.load(f)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
            with open(TOKEN_FILE, "wb") as f:
                pickle.dump(creds, f)
        else:
            raise RuntimeError("token.pickle 없거나 만료됨. Windows에서 복사 필요")
    return creds


def get_sheet():
    import gspread
    creds = get_credentials()
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(SPREADSHEET_ID)
    return sh.worksheet(SHEET_접수명단)


def to_int(val):
    """구글시트 숫자값 → int (빈값·문자 → 0)"""
    try:
        return int(str(val).replace(",", "").replace(" ", "") or 0)
    except (ValueError, TypeError):
        return 0


def match_and_update(amount: int, sender_name: str):
    """접수명단에서 입금자명 + 금액 매칭 → 입금체크 O 업데이트

    금액 검증: 할인가 또는 수수료와 일치해야 O 처리
    불일치 시 '금액불일치' 로그 → 수동 확인
    """
    ws = get_sheet()
    headers = ws.row_values(1)
    rows    = ws.get_all_records()

    if "입금체크" not in headers:
        print("  [경고] '입금체크' 컬럼 없음")
        return None

    col_입금체크 = headers.index("입금체크") + 1

    for i, row in enumerate(rows, start=2):
        name = str(row.get("성명", "") or "").strip()
        if name != sender_name:
            continue

        # 이름 매칭 → 금액 검증
        할인가 = to_int(row.get("할인가", 0))
        수수료  = to_int(row.get("수수료", 0))

        if amount == 할인가 or amount == 수수료:
            ws.update_cell(i, col_입금체크, "O")
            print(f"  [매칭] {name} / {amount:,}원 → 입금체크 O")
            return row
        else:
            print(f"  [금액불일치] {name} 입금 {amount:,}원 | 할인가 {할인가:,}원 / 수수료 {수수료:,}원 → 수동 확인")
            return None

    print(f"  [미매칭] '{sender_name}' 접수명단에 없음 → 수동 확인 필요")
    return None


# ===== 솔라피 자료접수 안내 발송 =====
def send_material_request(customer: dict):
    """솔라피 알림톡 - 자료접수 안내"""
    if not SOLAPI_KEY or not SOLAPI_TEMPLATE_자료접수:
        print("  [솔라피] 설정 미완료 → 발송 생략")
        return

    phone = str(customer.get("핸드폰번호", "") or "").strip()
    name  = str(customer.get("성명", "") or "").strip()
    if not phone:
        print(f"  [솔라피] 핸드폰번호 없음 → 발송 생략")
        return

    # 솔라피 API 호출 (알림톡)
    import hmac, hashlib, uuid
    date_str  = datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%S.000Z")
    salt      = str(uuid.uuid4()).replace("-", "")
    signature = hmac.new(
        SOLAPI_SECRET.encode(),
        f"{date_str}{salt}".encode(),
        hashlib.sha256,
    ).hexdigest()

    headers = {
        "Authorization": f"HMAC-SHA256 apiKey={SOLAPI_KEY}, date={date_str}, salt={salt}, signature={signature}",
        "Content-Type":  "application/json",
    }
    payload = {
        "message": {
            "to":          phone,
            "from":        SOLAPI_SENDER,
            "type":        "ATA",
            "kakaoOptions": {
                "pfId":       "",   # TODO: 플러스친구 ID
                "templateId": SOLAPI_TEMPLATE_자료접수,
                "variables":  {"#{고객명}": name},
            },
        }
    }
    resp = requests.post(
        "https://api.solapi.com/messages/v4/send",
        headers=headers,
        data=json.dumps(payload),
        timeout=10,
    )
    if resp.status_code == 200:
        print(f"  [솔라피] 자료접수 안내 발송 완료 → {name} ({phone})")
    else:
        print(f"  [솔라피] 발송 실패: {resp.status_code} {resp.text[:100]}")


# ===== 메인 루프 =====
def main():
    print(f"[카카오뱅크 입금 모니터] {datetime.now():%Y-%m-%d %H:%M:%S}")
    print(f"  DB: {DB_PATH}")
    print(f"  폴링 간격: {POLL_SEC}초")
    print(f"  Ctrl+C 로 종료\n")

    last_rowid = get_last_rowid()
    print(f"  시작 ROWID: {last_rowid}\n")

    while True:
        try:
            new_sms = get_new_kakao_sms(last_rowid)
            for rowid, msg, dt in new_sms:
                print(f"[카카오뱅크 SMS] {dt}")
                print(f"  내용: {msg[:150]}")

                result = parse_deposit(msg)
                if result:
                    amount, sender = result
                    print(f"  ★ 입금 감지: {sender} / {amount:,}원")
                    customer = match_and_update(amount, sender)
                    if customer:
                        send_material_request(customer)
                else:
                    print(f"  파싱 실패 → 수동 확인: {msg[:80]}")

                last_rowid = max(last_rowid, rowid)

        except KeyboardInterrupt:
            print("\n[종료]")
            break
        except Exception as e:
            print(f"  [에러] {e}")

        time.sleep(POLL_SEC)


if __name__ == "__main__":
    main()
