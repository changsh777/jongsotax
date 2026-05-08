"""페이지 레벨 CDP로 홈택스 DOM 검사"""
import sys, json, time, asyncio

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

import websockets

async def main():
    import requests
    tabs = requests.get("http://localhost:9222/json").json()
    print(f"탭 {len(tabs)}개")

    ht_tab = None
    for t in tabs:
        if "hometax.go.kr" in t.get("url",""):
            ht_tab = t
            break

    if not ht_tab:
        print("홈택스 탭 없음!")
        return

    ws_url = ht_tab["webSocketDebuggerUrl"]
    print(f"연결: {ws_url}")

    async with websockets.connect(ws_url) as ws:
        print("연결 성공!")

        async def send_cmd(method, params=None, cmd_id=1):
            msg = {"id": cmd_id, "method": method, "params": params or {}}
            await ws.send(json.dumps(msg))
            while True:
                resp = json.loads(await asyncio.wait_for(ws.recv(), timeout=15))
                if resp.get("id") == cmd_id:
                    return resp.get("result")

        # 버튼 목록 가져오기
        result = await send_cmd("Runtime.evaluate", {
            "expression": """(function() {
    var btns = Array.from(document.querySelectorAll('input[type=button], button'));
    return JSON.stringify(btns.map(function(el) {
        return {
            id: el.id || '',
            val: (el.value || el.innerText || '').trim().slice(0,40),
            cls: (el.className || '').slice(0,60),
            vis: el.offsetParent !== null
        };
    }));
})()""",
            "returnByValue": True
        })

        buttons = json.loads(result.get("result", {}).get("value", "[]"))
        print(f"\n버튼 {len(buttons)}개:")
        for b in buttons:
            if b.get('vis'):
                print(f"  [VISIBLE] id={b['id']:50s} val={b['val']:30s}")
            else:
                print(f"           id={b['id']:50s} val={b['val']:30s}")

        # select 박스
        result2 = await send_cmd("Runtime.evaluate", {
            "expression": """(function() {
    var sels = Array.from(document.querySelectorAll('select'));
    return JSON.stringify(sels.map(function(el) {
        return {
            id: el.id || '',
            opts: Array.from(el.options).map(function(o) { return o.value + ':' + o.text; }).join('|')
        };
    }));
})()""",
            "returnByValue": True
        }, cmd_id=2)

        selects = json.loads(result2.get("result", {}).get("value", "[]"))
        print(f"\nselect {len(selects)}개:")
        for s in selects:
            print(f"  id={s['id']:50s} opts={s['opts'][:80]}")

        # 빨강 배경 버튼
        result3 = await send_cmd("Runtime.evaluate", {
            "expression": """(function() {
    var all = Array.from(document.querySelectorAll('input, button, a'));
    var red = all.filter(function(el) {
        var bg = window.getComputedStyle(el).backgroundColor;
        var m = bg.match(/rgb\\((\\d+),\\s*(\\d+),\\s*(\\d+)\\)/);
        if (!m) return false;
        var r=+m[1], g=+m[2], b=+m[3];
        return (r>150 && g<120 && b<120) || (r>200 && g>80 && g<160 && b<80);
    });
    return JSON.stringify(red.map(function(el) {
        var bg = window.getComputedStyle(el).backgroundColor;
        return {
            tag: el.tagName,
            id: el.id || '',
            val: (el.value || el.innerText || '').trim().slice(0,50),
            bg: bg,
            cls: (el.className||'').slice(0,50)
        };
    }));
})()""",
            "returnByValue": True
        }, cmd_id=3)

        colored = json.loads(result3.get("result", {}).get("value", "[]"))
        print(f"\n컬러 버튼 {len(colored)}개:")
        for b in colored:
            print(f"  {b['tag']:8s} id={b['id'][:50]:50s} val={b['val'][:30]:30s} bg={b['bg']}")

asyncio.run(main())
