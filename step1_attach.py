"""
1단계 (CDP 방식): 이미 떠있는 크롬에 붙어서 신고도움서비스 진입
- launch_chrome.bat로 크롬을 먼저 띄워서 로그인해둬야 함
- 이 스크립트는 그 크롬에 9222 포트로 붙어서 자동화만 담당
"""
from playwright.sync_api import sync_playwright

REPORT_HELP_URL = (
    "https://hometax.go.kr/websquare/websquare.html"
    "?w2xPath=/ui/pp/index_pp.xml"
    "&tmIdx=06&tm2lIdx=0601000000&tm3lIdx=0601200000"
)


def main():
    with sync_playwright() as p:
        print("[1] 9222 포트로 크롬에 붙습니다.")
        browser = p.chromium.connect_over_cdp("http://localhost:9222")

        ctx = browser.contexts[0]
        page = ctx.pages[0] if ctx.pages else ctx.new_page()
        page.bring_to_front()

        print(f"[2] 현재 URL: {page.url}")
        print("[3] 신고도움서비스로 이동합니다.")
        page.goto(REPORT_HELP_URL, wait_until="domcontentloaded")

        print("[4] 진입 완료. 화면 탐색하세요.")
        print("    종료하려면 터미널에서 ENTER.")
        input()


if __name__ == "__main__":
    main()
