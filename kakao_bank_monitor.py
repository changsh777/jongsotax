"""
kakao_bank_monitor.py - 카카오뱅크 입금 알림 실시간 감지

방식: Windows 토스트 알림 리스너
- 카카오톡으로 오는 카카오뱅크 입금 알림 캡처
- 입금자명 + 금액 파싱
- 구글시트 접수명단에서 매칭 → 입금체크 업데이트
- 다음 단계: 자료접수 안내 메시지 생성 (솔라피 발송용)

실행:
  python kakao_bank_monitor.py          # 모니터링 시작
  python kakao_bank_monitor.py --test   # 알림 파싱 테스트 (더미 메시지)

카카오뱅크 입금 알림 형식 (예시):
  "[카카오뱅크] 입금 300,000원\n잔액 1,234,567원\n입금자명: 홍길동"
  또는
  "입금 300,000원 홍길동"  (짧은 형식)
"""
import asyncio
import re
import sys
import argparse
from datetime import datetime

sys.path.insert(0, r"F:\종소세2026")

# ===== 입금 메시지 파싱 =====
# 카카오뱅크 알림 패턴 (실제 메시지 확인 후 수정 가능)
PATTERNS = [
    # "[카카오뱅크] 입금 300,000원 ... 입금자명: 홍길동"
    re.compile(r"입금\s+([\d,]+)원.*?입금자명[:\s]+([^\n\r]+)", re.DOTALL),
    # "입금 300,000원 홍길동" (짧은 형식)
    re.compile(r"입금\s+([\d,]+)원\s+(.+)"),
]


def parse_deposit(text: str):
    """입금 알림 텍스트 → (금액, 입금자명) 또는 None"""
    if not text or "입금" not in text:
        return None
    for pat in PATTERNS:
        m = pat.search(text)
        if m:
            amount_str = m.group(1).replace(",", "")
            name = m.group(2).strip()
            try:
                amount = int(amount_str)
                return amount, name
            except ValueError:
                continue
    return None


# ===== 구글시트 매칭 =====
def match_and_update(amount: int, sender_name: str):
    """접수명단에서 입금자명 매칭 → 입금체크"""
    from gsheet_writer import get_credentials
    import gspread

    creds = get_credentials()
    gc = gspread.authorize(creds)
    sh = gc.open_by_key("1oh31k00Oa2lZWvu5fnBRVmurdlll1YEG8Fefi5FRfBI")
    ws = sh.worksheet("접수명단")
    rows = ws.get_all_records()

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    matched = False

    for i, row in enumerate(rows, start=2):
        name = str(row.get("성명", "") or "").strip()
        if name == sender_name:
            # 입금체크 컬럼 위치 찾기
            headers = ws.row_values(1)
            if "입금체크" in headers:
                col = headers.index("입금체크") + 1
                ws.update_cell(i, col, "O")
            print(f"  [매칭] {name} 입금 {amount:,}원 → 입금체크 완료")
            matched = True
            return row  # 다음 단계 처리용

    if not matched:
        print(f"  [미매칭] '{sender_name}' 접수명단에 없음 (수동 확인 필요)")
    return None


# ===== 실시간 알림 리스너 =====
async def listen_notifications():
    from winsdk.windows.ui.notifications.management import (
        UserNotificationListener,
        UserNotificationListenerAccessStatus,
    )

    listener = UserNotificationListener.current
    status = await listener.request_access_async()

    if status != UserNotificationListenerAccessStatus.ALLOWED:
        print("[에러] 알림 접근 권한 없음")
        print("  설정 → 알림 및 작업 → '앱과 다른 발신자의 알림 받기' 켜기")
        return

    print(f"[카카오뱅크 입금 모니터 시작] {datetime.now():%Y-%m-%d %H:%M:%S}")
    print("  카카오뱅크 입금 알림 대기 중... (Ctrl+C로 종료)\n")

    def on_notification_added(sender, args):
        try:
            notif = args.notification
            app_name = ""
            try:
                app_name = notif.app_info.display_info.display_name
            except Exception:
                pass

            # 카카오톡 알림만 처리
            if "카카오" not in app_name and "Kakao" not in app_name:
                return

            # 알림 텍스트 추출
            content = notif.notification.content
            xml = content.get_xml()

            # XML에서 텍스트 추출 (간단 정규식)
            texts = re.findall(r"<text[^>]*>([^<]+)</text>", xml)
            full_text = "\n".join(texts)

            print(f"[카카오톡 알림] {datetime.now():%H:%M:%S}")
            print(f"  내용: {full_text[:200]}")

            # 카카오뱅크 입금 여부 확인
            if "카카오뱅크" in full_text and "입금" in full_text:
                result = parse_deposit(full_text)
                if result:
                    amount, sender = result
                    print(f"  ★ 입금 감지: {sender} / {amount:,}원")
                    matched_row = match_and_update(amount, sender)
                    if matched_row:
                        print(f"  → 자료접수 안내 대기 (솔라피 발송 준비)")
                        # TODO: 솔라피 자료접수 안내 발송
                else:
                    print(f"  카카오뱅크 알림이나 파싱 실패 → 수동 확인")
                    print(f"  전문: {full_text}")

        except Exception as e:
            print(f"  [알림 처리 에러] {e}")

    # 이벤트 핸들러 등록
    token = listener.add_notification_changed(on_notification_added)

    try:
        # 무한 대기
        while True:
            await asyncio.sleep(1)
    except KeyboardInterrupt:
        print("\n[종료]")
    finally:
        listener.remove_notification_changed(token)


# ===== 테스트 =====
def test_parse():
    """다양한 카카오뱅크 메시지 파싱 테스트"""
    samples = [
        "[카카오뱅크] 입금 300,000원\n잔액 1,234,567원\n입금자명: 홍길동",
        "카카오뱅크\n입금 150,000원 홍길동",
        "[카카오뱅크]\n입금 80,000원\n입금자명: 박수경",
        "카카오뱅크 입금 200,000원\n입금자명: 이영희",
    ]
    print("[파싱 테스트]")
    for s in samples:
        result = parse_deposit(s)
        print(f"  입력: {repr(s[:50])}")
        print(f"  결과: {result}")
        print()


# ===== 메인 =====
def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--test", action="store_true", help="파싱 테스트만")
    args = parser.parse_args()

    if args.test:
        test_parse()
        return

    asyncio.run(listen_notifications())


if __name__ == "__main__":
    main()
