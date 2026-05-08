"""window.open 가로채기 v2 — about:blank 후 location.href 패턴 대응"""
import requests, json, asyncio, websockets

CDP = "http://localhost:9222"

async def _eval(ws, code, cmd_id=1):
    await ws.send(json.dumps({"id": cmd_id, "method": "Runtime.evaluate",
                              "params": {"expression": code, "returnByValue": True}}))
    while True:
        r = json.loads(await asyncio.wait_for(ws.recv(), timeout=15))
        if r.get("id") == cmd_id:
            v = r.get("result", {}).get("result", {})
            return v.get("value")

async def run():
    tabs = requests.get(f"{CDP}/json").json()
    main = next((t for t in tabs
                 if "hometax.go.kr" in t.get("url","")
                 and "websquare.html" in t.get("url","")
                 and "sesw." not in t.get("url","")), None)
    if not main:
        print("홈택스 탭 없음"); return
    print(f"탭: {main['url'][:80]}")

    async with websockets.connect(main["webSocketDebuggerUrl"], ping_interval=None) as ws:
        # 1. 가짜 window 객체로 URL 캡처 (about:blank → location.href 설정 패턴 대응)
        r2 = await _eval(ws, """(function(){
    window._popup_url = null;
    window._popup_html = null;
    window.open = function(url, name, features) {
        var captured = url || '';
        var fakeWin = {
            _url: captured,
            location: {
                get href(){ return window._popup_url || captured; },
                set href(v){ window._popup_url = v; console.log('[POPUP_URL]', v); }
            },
            document: {
                write: function(html){ window._popup_html = (html||'').slice(0,500); },
                writeln: function(html){ window._popup_html = (html||'').slice(0,500); },
                close: function(){}
            },
            focus: function(){}, blur: function(){}, close: function(){}
        };
        window._popup_url = captured;
        console.log('[OPEN]', captured);
        return fakeWin;
    };
    return 'intercept_v2_ok';
})()""", cmd_id=2)
        print(f"가로채기: {r2}")

        # 2. 버튼 클릭
        r3 = await _eval(ws, """(function(){
    var rows = Array.from(document.querySelectorAll('[id*="UTERNAAZ0Z31_wframe"] table tbody tr'))
        .filter(function(tr){ return tr.querySelectorAll('td').length >= 13; });
    var cell = rows[0] && rows[0].querySelectorAll('td')[10];
    var btn = cell && cell.querySelector('a, input[type=button], button');
    if (!btn) return 'no_btn';
    btn.click();
    return 'clicked:' + (btn.innerText||btn.value||'').trim().slice(0,30);
})()""", cmd_id=3)
        print(f"클릭: {r3}")

        # 3. 폴링 — location.href 변경 대기 (최대 3초)
        final_url = None
        for i in range(6):
            await asyncio.sleep(0.5)
            u = await _eval(ws, "window._popup_url", cmd_id=10+i)
            h = await _eval(ws, "window._popup_html ? window._popup_html.slice(0,200) : null", cmd_id=20+i)
            print(f"  [{i*0.5:.1f}s] URL={u}  HTML={h}")
            if u and u != "about:blank":
                final_url = u
                break
            if h:
                print(f"  → document.write 감지!")
                break

        print(f"\n최종 팝업 URL: {final_url}")
        if final_url:
            print("→ 이 URL로 CDP Target.createTarget 직접 열기 가능!")
        elif r3 and "clicked" in str(r3):
            print("→ about:blank 유지 — document.write 방식이거나 다른 navigate 방식")
            print(f"  HTML snippet: {await _eval(ws, 'window._popup_html', cmd_id=99)}")

asyncio.run(run())
