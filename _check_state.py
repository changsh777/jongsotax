"""현재 페이지 상태 + 세션 확인"""
import sys, asyncio, json, requests
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')
import websockets

async def main():
    tabs = requests.get("http://localhost:9222/json").json()
    print(f"탭 {len(tabs)}개:")
    for t in tabs:
        print(f"  {t.get('type','?')} | {t.get('url','')[:100]}")

    ht = next((t for t in tabs if "hometax.go.kr" in t.get("url","")), None)
    if not ht:
        print("홈택스 탭 없음!"); return

    async with websockets.connect(ht["webSocketDebuggerUrl"]) as ws:
        async def ev(code, cid=1):
            await ws.send(json.dumps({"id":cid,"method":"Runtime.evaluate",
                "params":{"expression":code,"returnByValue":True}}))
            while True:
                r = json.loads(await asyncio.wait_for(ws.recv(), timeout=15))
                if r.get("id")==cid:
                    return r.get("result",{}).get("result",{}).get("value")

        url = await ev("window.location.href", 1)
        print(f"\n현재 URL: {url}")

        # 로그인 상태 확인
        login_check = await ev("""(function(){
    var logoutBtn = document.querySelector('#mf_wfHeader_btnTopLgn, [id*="btnLogout"]');
    var loginBtn = document.querySelector('[id*="btnTopLgn"]');
    return JSON.stringify({
        logoutBtn: logoutBtn ? logoutBtn.value || logoutBtn.innerText : null,
        loginBtn: loginBtn ? loginBtn.value || loginBtn.innerText : null
    });
})()""", 2)
        print(f"로그인 상태: {login_check}")

        # btnRtnInqr 존재 확인
        btn_check = await ev("""(function(){
    var btn = document.getElementById('mf_txppWframe_btnRtnInqr');
    if (!btn) return 'NOT FOUND';
    return {id:btn.id, val:btn.value||'', vis:(btn.offsetParent!==null)};
})()""", 3)
        print(f"btnRtnInqr: {btn_check}")

asyncio.run(main())
