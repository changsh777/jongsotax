"""
접수증 버튼 클릭 후 DOM 변화 추적
같은 페이지 내 팝업/iframe이 열리는지 확인
"""
import sys, asyncio, json, requests
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')
import websockets

async def main():
    tabs = requests.get("http://localhost:9222/json").json()
    ht = next((t for t in tabs if "hometax.go.kr" in t.get("url","")), None)
    if not ht:
        print("홈택스 탭 없음!"); return

    async with websockets.connect(ht["webSocketDebuggerUrl"]) as ws:
        cmd_id = [0]
        async def ev(code):
            cmd_id[0] += 1
            cid = cmd_id[0]
            await ws.send(json.dumps({"id":cid,"method":"Runtime.evaluate",
                "params":{"expression":code,"returnByValue":True}}))
            while True:
                r = json.loads(await asyncio.wait_for(ws.recv(), timeout=30))
                if r.get("id")==cid:
                    v = r.get("result",{}).get("result",{}).get("value")
                    return json.loads(v) if isinstance(v,str) and v.startswith('[') else v

        # 현재 상태 스냅샷
        before = await ev("""(function(){
    var iframes = Array.from(document.querySelectorAll('iframe')).map(function(f){
        return {id:f.id||'', src:f.src||'', name:f.name||''};
    });
    var visibleDivs = Array.from(document.querySelectorAll('div[style*="display:block"], div[style*="visibility:visible"]'))
        .filter(function(d){ return d.offsetWidth > 200 && d.offsetHeight > 100; })
        .map(function(d){ return {id:d.id||'', cls:(d.className||'').slice(0,40)}; });
    return JSON.stringify({iframes:iframes.length, visibleDivs:visibleDivs.slice(0,5)});
})()""")
        print(f"클릭 전: {before}")

        # 첫 번째 데이터 행의 접수증(col[12]) 버튼 클릭
        click_r = await ev("""(function(){
    var container = document.querySelector('[id*="UTERNAAZ0Z31_wframe"]');
    if (!container) return 'no_container';
    var rows = Array.from(container.querySelectorAll('table tbody tr'))
        .filter(function(tr){ return tr.querySelectorAll('td').length >= 13; });
    if (!rows.length) return 'no_rows (팝업 닫혀있음?)';
    var row = rows[0];
    var tds = Array.from(row.querySelectorAll('td'));
    var btn = tds[12] ? tds[12].querySelector('input[type=button], button') : null;
    if (!btn) return 'no_btn_col12';
    var name = tds[6] ? tds[6].innerText.trim() : '?';
    btn.click();
    return 'clicked:' + name + ':' + (btn.value||'?');
})()""")
        print(f"클릭: {click_r}")

        # 0.5초 후 변화 확인
        await asyncio.sleep(0.5)
        after05 = await ev("""(function(){
    var iframes = Array.from(document.querySelectorAll('iframe')).map(function(f){
        return {id:f.id||'', src:(f.src||f.getAttribute('src')||'').slice(0,80), name:f.name||''};
    });
    return JSON.stringify({iframes:iframes});
})()""")
        print(f"\n0.5초 후 iframe: {after05}")

        # 1초 후
        await asyncio.sleep(0.5)
        after1 = await ev("""(function(){
    var iframes = Array.from(document.querySelectorAll('iframe')).map(function(f){
        return {id:f.id||'', src:(f.src||f.contentWindow&&f.contentWindow.location.href||'').slice(0,80)};
    });
    // 새로 생긴 큰 div
    var bigDivs = Array.from(document.querySelectorAll('div'))
        .filter(function(d){
            var r = d.getBoundingClientRect();
            return r.width > 300 && r.height > 200 && d.id;
        })
        .map(function(d){ return {id:d.id.slice(0,60), w:Math.round(d.getBoundingClientRect().width)}; });
    return JSON.stringify({iframes:iframes, bigDivs:bigDivs.slice(0,10)});
})()""")
        print(f"1초 후: {after1}")

        # 3초 후 - 새 팝업 요소 탐색
        await asyncio.sleep(2)
        after3 = await ev("""(function(){
    // 접수증/clipreport 관련 요소 탐색
    var keys = ['접수증', 'rcpt', 'receipt', 'clipreport', 'clip', 'report'];
    var found = [];
    keys.forEach(function(k){
        var els = Array.from(document.querySelectorAll('[id*="'+k+'"], [class*="'+k+'"], iframe[src*="'+k+'"]'));
        els.forEach(function(el){
            if (el.offsetParent) {
                found.push({k:k, tag:el.tagName, id:el.id||'', src:(el.src||'').slice(0,60)});
            }
        });
    });

    // window.open이 호출됐는지 확인 (monkey-patch)
    var iframes = Array.from(document.querySelectorAll('iframe')).map(function(f){
        var src = '';
        try { src = f.contentWindow.location.href; } catch(e) { src = f.src || '?'; }
        return {id:f.id, src:src.slice(0,80)};
    });

    // 현재 URL이 변경됐는지
    return JSON.stringify({
        found:found,
        iframes:iframes,
        currentUrl: window.location.href.slice(0,80)
    });
})()""")
        print(f"\n3초 후 접수증 관련 요소:\n{after3}")

        # window.open 감시 (미래 클릭용)
        await ev("""(function(){
    if (!window._openWatched) {
        window._openWatched = true;
        window._openCalls = [];
        var orig = window.open;
        window.open = function(url, name, features) {
            window._openCalls.push({url:url, name:name});
            console.log('window.open called:', url, name);
            return orig.apply(this, arguments);
        };
    }
    return 'watching';
})()""")
        print("\nwindow.open 감시 설치 완료")

        # 다시 클릭 (감시 후)
        click_r2 = await ev("""(function(){
    var container = document.querySelector('[id*="UTERNAAZ0Z31_wframe"]');
    if (!container) return 'no_container';
    var rows = Array.from(container.querySelectorAll('table tbody tr'))
        .filter(function(tr){ return tr.querySelectorAll('td').length >= 13; });
    if (!rows.length) return 'no_rows';
    var tds = Array.from(rows[0].querySelectorAll('td'));
    var btn = tds[12] ? tds[12].querySelector('input[type=button], button') : null;
    if (!btn) return 'no_btn';
    btn.click();
    return 'clicked2';
})()""")
        print(f"2차 클릭: {click_r2}")

        await asyncio.sleep(2)

        open_calls = await ev("JSON.stringify(window._openCalls||[])")
        print(f"\nwindow.open 호출 기록: {open_calls}")

asyncio.run(main())
