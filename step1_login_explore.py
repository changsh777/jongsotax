"""
1단계: 홈택스 로그인 + 신고도움서비스 진입 확인
- 브라우저 띄우고 사용자가 직접 공동인증서로 로그인
- ENTER 입력하면 신고도움서비스 URL로 이동
- 브라우저 열어둔 채로 화면 구조 탐색
"""
from playwright.sync_api import sync_playwright
from pathlib import Path

USER_DATA_DIR = Path(r"F:\종소세2026\.browser_profile")
HOMETAX_MAIN = "https://hometax.go.kr"
REPORT_HELP_URL = (
    "https://hometax.go.kr/websquare/websquare.html"
    "?w2xPath=/ui/pp/index_pp.xml"
    "&tmIdx=06&tm2lIdx=0601000000&tm3lIdx=0601200000"
)


def main():
    USER_DATA_DIR.mkdir(parents=True, exist_ok=True)

    with sync_playwright() as p:
        ctx = p.chromium.launch_persistent_context(
            user_data_dir=str(USER_DATA_DIR),
            channel="chrome",
            headless=False,
            viewport={"width": 1400, "height": 900},
            accept_downloads=True,
            args=["--disable-blink-features=AutomationControlled"],
        )
        page = ctx.pages[0] if ctx.pages else ctx.new_page()

        print("[1] 홈택스 메인으로 이동합니다.")
        page.goto(HOMETAX_MAIN, wait_until="domcontentloaded")

        input("[2] 로그인 완료되면 ENTER 눌러주세요 (공동인증서로 직접 로그인)... ")

        print("[3] 신고도움서비스로 이동합니다.")
        page.goto(REPORT_HELP_URL, wait_until="domcontentloaded")

        print("[4] 브라우저 열어둔 상태. 화면 탐색하세요.")
        print("    종료하려면 이 터미널에서 ENTER 한 번 더.")
        input()

        ctx.close()


if __name__ == "__main__":
    main()
