"""
접수증 팝업 window 레퍼런스 직접 캡처
window.open의 반환값(팝업 window 객체)을 모니터링
"""
import sys, asyncio, json, requests, time
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
        async def ev(code, timeout=30):
            cmd_id[0] += 1
            cid = cmd_id[0]
            await ws.send(json.dumps({"id":cid,"method":"Runtime.evaluate",
                "params":{"expression":code,"returnByValue":True}}))
            while True:
                r = json.loads(await asyncio.wait_for(ws.recv(), timeout=timeout))
                if r.get("id")==cid:
                    v = r.get("result",{}).get("result",{})
                    val = v.get("value")
                    return json.loads(val) if isinstance(val,str) and (val.startswith('{') or val.startswith('[')) else val

        # 1. 팝업 레퍼런스 캡처용 monkey-patch 설치
        setup = await ev("""(function(){
    if (window._popupCapture) return 'already_setup';
    window._popupCapture = {};
    var origOpen = window.open;
    window.open = function(url, name, features) {
        var w = origOpen.apply(this, arguments);
        var info = {url:url||'', name:name||'', blocked: (w===null)};
        window._popupCapture[name||('_'+Date.now())] = {ref:w, info:info};
        console.log('[popup-capture] open called:', JSON.stringify(info));
        return w;
    };
    return 'setup_ok';
})()""")
        print(f"팝업 캡처 설치: {setup}")

        # 2. 신고내역 팝업 열기 (없으면)
        open_popup = await ev("""(function(){
    var btn = document.getElementById('mf_txppWframe_btnRtnInqr');
    if (!btn) return 'no_btn';
    var container = document.querySelector('[id*="UTERNAAZ0Z31_wframe"]');
    // 팝업이 이미 열려있는지 확인
    if (container) {
        var rows = Array.from(container.querySelectorAll('table tbody tr'))
            .filter(function(tr){ return tr.querySelectorAll('td').length >= 13; });
        if (rows.length > 0) return 'already_open:' + rows.length + '행';
    }
    btn.click();
    return 'opened';
})()""")
        print(f"신고내역 팝업: {open_popup}")
        if 'opened' in str(open_popup):
            await asyncio.sleep(3)
            # 조회 버튼
            await ev("""(function(){
    var btn = document.getElementById('mf_txppWframe_UTERNAAZ0Z31_wframe_trigger70_UTERNAAZ31');
    if (btn) btn.click();
})()""")
            await asyncio.sleep(4)

        # 3. 접수증 버튼 클릭
        click = await ev("""(function(){
    var container = document.querySelector('[id*="UTERNAAZ0Z31_wframe"]');
    if (!container) return 'no_container';
    var rows = Array.from(container.querySelectorAll('table tbody tr'))
        .filter(function(tr){ return tr.querySelectorAll('td').length >= 13; });
    if (!rows.length) return 'no_rows';
    var tds = Array.from(rows[0].querySelectorAll('td'));
    var name = tds[6] ? tds[6].innerText.trim() : '?';
    var btn = tds[12] ? tds[12].querySelector('input[type=button], button') : null;
    if (!btn) return 'no_btn:col12';
    window._popupCapture = {};  // 초기화
    btn.click();
    return 'clicked:' + name;
})()""")
        print(f"\n접수증 버튼 클릭: {click}")

        # 4. 팝업 레퍼런스 확인 (최대 5초)
        for i in range(10):
            await asyncio.sleep(0.5)
            status = await ev("""(function(){
    var caps = window._popupCapture || {};
    var results = [];
    Object.keys(caps).forEach(function(k) {
        var cap = caps[k];
        var url = '';
        var closed = true;
        if (cap.ref) {
            closed = cap.ref.closed;
            if (!closed) {
                try { url = cap.ref.location.href; } catch(e) { url = 'cross-origin(' + (cap.info.url||'?') + ')'; }
            }
        }
        results.push({
            name: k,
            blocked: cap.info.blocked,
            closed: closed,
            url: url || cap.info.url
        });
    });
    return JSON.stringify(results);
})()""")
            if status and status != '[]':
                print(f"  {i*0.5:.1f}초: 팝업 상태 = {status}")
                # URL이 about:blank가 아니면 성공
                try:
                    items = json.loads(status) if isinstance(status, str) else status
                    for item in items:
                        if item.get('url','') not in ('about:blank', '', 'about:blank'):
                            print(f"\n  *** 접수증 URL 발견: {item['url']} ***")
                except Exception:
                    pass
            else:
                print(f"  {i*0.5:.1f}초: 팝업 캡처 없음")

        # 5. 최종 상태
        final = await ev("""(function(){
    var caps = window._popupCapture || {};
    var results = [];
    Object.keys(caps).forEach(function(k) {
        var cap = caps[k];
        var info = {name:k, blocked:cap.info.blocked};
        if (cap.ref && !cap.ref.closed) {
            try {
                info.url = cap.ref.location.href;
                // 팝업 내 HTML 스니펫
                info.html = cap.ref.document.documentElement.innerHTML.slice(0,200);
            } catch(e) {
                info.url = 'cross-origin';
                info.error = e.toString();
            }
        } else {
            info.closed = true;
        }
        results.push(info);
    });
    return JSON.stringify(results);
})()""")
        print(f"\n최종 팝업 상태:\n{final}")

asyncio.run(main())
