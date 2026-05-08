"""
접수증 버튼 클릭 후 탭/팝업 모니터링
실행 후 직접 Edge에서 접수증 보기 버튼 눌러보세요
"""
import sys, asyncio, requests, time, json
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')
import websockets

ENDPOINTS = [
    "http://localhost:9222/json",
    "http://localhost:9222/json/list",
]

def get_all_targets():
    targets = {}
    for url in ENDPOINTS:
        try:
            data = requests.get(url, timeout=3).json()
            if isinstance(data, list):
                for t in data:
                    targets[t.get("id","")] = t
        except Exception:
            pass
    return targets

before = get_all_targets()
print(f"모니터링 시작. 현재 탭/타겟 {len(before)}개:")
for tid, t in before.items():
    print(f"  {t.get('type','?'):8s} {t.get('url','')[:70]}")

print("\n지금 Edge에서 신고내역 조회 팝업의 접수증 '보기' 버튼을 클릭하세요...")
print("(30초간 모니터링)")

for i in range(60):
    time.sleep(0.5)
    after = get_all_targets()
    new_ones = {k: v for k, v in after.items() if k not in before}
    if new_ones:
        print(f"\n!!! {i*0.5:.1f}초에 새 타겟 발견 !!!")
        for tid, t in new_ones.items():
            print(f"  type={t.get('type','?')} url={t.get('url','')[:80]}")
            print(f"  wsDebugger={t.get('webSocketDebuggerUrl','')[:60]}")
        before.update(new_ones)
    elif i % 6 == 0:
        print(f"  {i*0.5:.0f}초 대기 중... (총 {len(after)}개 탭)")

print("\n모니터링 완료")
