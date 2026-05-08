"""
현재 열린 ClipReport 탭에서 PDF 즉시 저장
"""
import sys, asyncio, json, requests, time
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')
import websockets
import os
os.environ.setdefault("SEOTAX_ENV", "nas")
sys.path.insert(0, str(__import__("pathlib").Path(__file__).parent))
from config import CUSTOMER_DIR

async def main():
    tabs = requests.get("http://localhost:9222/json").json()
    print("현재 탭:")
    for t in tabs:
        print(f"  {t.get('url','')[:80]}")

    # ClipReport 탭 찾기
    cr_tab = next((t for t in tabs if "clipreport" in t.get("url","").lower()), None)
    if not cr_tab:
        print("ClipReport 탭 없음!")
        return

    print(f"\nClipReport 탭 연결: {cr_tab['url']}")

    async with websockets.connect(cr_tab["webSocketDebuggerUrl"]) as ws:
        async def ev(code, cid=1):
            await ws.send(json.dumps({"id":cid,"method":"Runtime.evaluate",
                "params":{"expression":code,"returnByValue":True}}))
            while True:
                r = json.loads(await asyncio.wait_for(ws.recv(), timeout=20))
                if r.get("id")==cid:
                    return r.get("result",{}).get("result",{}).get("value")

        # 페이지 로딩 확인
        state = await ev("document.readyState", 1)
        print(f"readyState: {state}")

        # 고객명 확인 (ClipReport 내용에서)
        title = await ev("document.title || document.querySelector('title')?.innerText || ''", 2)
        print(f"타이틀: {title}")

        body_text = await ev("document.body ? document.body.innerText.slice(0,300) : 'no body'", 3)
        print(f"내용 미리보기:\n{body_text}")

        # PDF 버튼 확인
        pdf_btn = await ev("""(function(){
    var btn = document.querySelector('.report_menu_pdf_button');
    if (!btn) return 'no pdf btn';
    return {
        cls: btn.className,
        disabled: btn.disabled,
        vis: btn.offsetParent !== null
    };
})()""", 4)
        print(f"\nPDF 버튼: {pdf_btn}")

        # 다운로드 경로 설정 (NAS 첫 번째 고객 폴더 찾기)
        # 일단 임시로 홈 디렉토리에 저장
        import pathlib
        save_dir = pathlib.Path.home() / "Downloads"
        save_dir.mkdir(exist_ok=True)

        # 다운로드 경로 설정
        await ws.send(json.dumps({
            "id": 5, "method": "Page.setDownloadBehavior",
            "params": {"behavior": "allow", "downloadPath": str(save_dir)}
        }))
        while True:
            r = json.loads(await asyncio.wait_for(ws.recv(), timeout=10))
            if r.get("id") == 5:
                print(f"다운로드 경로 설정: {r}")
                break

        await asyncio.sleep(2)

        # PDF 다운로드 클릭
        clicked = await ev("""(function(){
    var btn = document.querySelector('.report_menu_pdf_button');
    if (!btn) return 'no btn';
    btn.classList.remove('report_menu_pdf_button_svg_dis');
    btn.classList.add('report_menu_pdf_button_svg');
    btn.disabled = false;
    btn.click();
    return 'clicked';
})()""", 6)
        print(f"PDF 클릭: {clicked}")

        # 다운로드 대기 (Page.downloadProgress 이벤트)
        await ws.send(json.dumps({"id": 7, "method": "Page.enable", "params": {}}))
        while True:
            r = json.loads(await asyncio.wait_for(ws.recv(), timeout=5))
            if r.get("id") == 7:
                break

        print("다운로드 대기 중...")
        for _ in range(60):
            await asyncio.sleep(0.5)
            try:
                msg = json.loads(await asyncio.wait_for(ws.recv(), timeout=0.2))
                method = msg.get("method","")
                if method == "Page.downloadWillBegin":
                    print(f"  다운로드 시작: {msg.get('params',{}).get('suggestedFilename','?')}")
                elif method == "Page.downloadProgress":
                    p = msg.get("params",{})
                    if p.get("state") == "completed":
                        print(f"  다운로드 완료!")
                        break
                    elif p.get("state") == "canceled":
                        print("  다운로드 취소됨")
                        break
            except asyncio.TimeoutError:
                pass

        # 저장된 파일 확인
        await asyncio.sleep(1)
        pdfs = sorted(save_dir.glob("*.pdf"), key=lambda f: f.stat().st_mtime, reverse=True)
        if pdfs:
            print(f"\n저장된 PDF: {pdfs[0]}")
        else:
            print("\nPDF 파일 없음")

asyncio.run(main())
