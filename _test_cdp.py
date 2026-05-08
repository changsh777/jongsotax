"""최소 CDP 연결 테스트"""
import sys, asyncio
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

async def main():
    from playwright.async_api import async_playwright
    print("시작...")
    async with async_playwright() as p:
        print("Playwright 초기화 완료")
        print("CDP 연결 중...")
        try:
            browser = await asyncio.wait_for(
                p.chromium.connect_over_cdp("http://localhost:9222", slow_mo=0),
                timeout=30
            )
            print(f"연결 성공! contexts: {len(browser.contexts)}")
            if browser.contexts:
                ctx = browser.contexts[0]
                print(f"pages: {len(ctx.pages)}")
        except asyncio.TimeoutError:
            print("30초 타임아웃!")
        except Exception as e:
            print(f"오류: {e}")
    print("완료")

asyncio.run(main())
