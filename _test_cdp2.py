"""playwright async 연결 테스트 - 타임아웃 단계별로"""
import sys, asyncio, time
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

async def main():
    from playwright.async_api import async_playwright

    print(f"시작: {time.strftime('%H:%M:%S')}")
    async with async_playwright() as p:
        print(f"PW 준비: {time.strftime('%H:%M:%S')}")

        # connect_over_cdp는 내부적으로 아래 작업을 함:
        # 1. HTTP GET /json/version → webSocketDebuggerUrl 취득
        # 2. WS 연결
        # 3. Target.getTargets
        # 4. 각 Target.attachToTarget
        # 타임아웃 없이 연결 시도

        try:
            print(f"connect 시작: {time.strftime('%H:%M:%S')}")
            browser = await p.chromium.connect_over_cdp(
                "http://localhost:9222",
                timeout=120000  # 명시적 2분
            )
            print(f"connect 완료: {time.strftime('%H:%M:%S')}")
            print(f"contexts: {len(browser.contexts)}")
        except Exception as e:
            print(f"오류 {time.strftime('%H:%M:%S')}: {e}")
    print(f"종료: {time.strftime('%H:%M:%S')}")

asyncio.run(main())
