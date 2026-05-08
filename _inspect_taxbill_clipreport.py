"""납부서 버튼 클릭 후 ClipReport 탭 HTML 구조 검사"""
import asyncio, json, requests, sys
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')
import websockets

CDP_PORT = 9222
COL_TAX = 13
SELECT_ROWNUM = "mf_txppWframe_UTERNAAZ0Z31_wframe_edtGrdRowNum"

def get_tabs():
    return requests.get(f'http://localhost:{CDP_PORT}/json').json()

def get_ht_tab():
    tabs = get_tabs()
    return next((t for t in tabs if 'websquare.html' in t.get('url','') and 'sesw.' not in t.get('url','')), None)

async def ev(ws, code, cid=1):
    await ws.send(json.dumps({'id':cid,'method':'Runtime.evaluate',
        'params':{'expression':code,'returnByValue':True}}))
    while True:
        r = json.loads(await asyncio.wait_for(ws.recv(), timeout=30))
        if r.get('id')==cid:
            return r.get('result',{}).get('result',{}).get('value')

async def main():
    ht = get_ht_tab()
    if not ht:
        print("홈택스 탭 없음")
        return
    print(f"홈택스 탭: {ht['url'][:60]}")

    known_ids = {t['id'] for t in get_tabs()}

    async with websockets.connect(ht['webSocketDebuggerUrl'], ping_interval=None) as ws:
        # 팝업 열려있는지 확인
        state = await ev(ws, f'document.getElementById("{SELECT_ROWNUM}") ? "open" : "closed"', 1)
        print(f"팝업 상태: {state}")
        if state != 'open':
            print("팝업을 먼저 열어주세요 (신고내역조회)")
            return

        # 지성환(row 7, 0-indexed) 납부서 버튼 클릭
        # 먼저 몇 번째 행인지 확인
        rows_info = await ev(ws, """(function(){
    var container = document.querySelector('[id*="UTERNAAZ0Z31_wframe"]');
    if (!container) return 'no_container';
    var rows = Array.from(container.querySelectorAll('table tbody tr'))
        .filter(function(tr){ return tr.querySelectorAll('td').length >= 13; });
    return rows.map(function(tr, idx){
        var tds = Array.from(tr.querySelectorAll('td'));
        var name = tds[6] ? tds[6].innerText.trim() : '';
        var jumin = tds[7] ? tds[7].innerText.trim().slice(0,6) : '';
        // 납부서 버튼 확인
        var taxCell = tds[13];
        var taxBtn = taxCell ? taxCell.querySelector('input[type=button],button') : null;
        var taxTxt = taxCell ? taxCell.innerText.trim() : '';
        return idx + ':' + name + '(' + jumin + ') tax=[' + taxTxt + '] btn=' + (taxBtn ? taxBtn.value||'?' : 'none');
    }).join('\n');
})()""", 2)
        print("=== 행 목록 ===")
        print(rows_info)

        # 지성환 행 찾기
        target_row = None
        if rows_info and '\n' in str(rows_info):
            for line in rows_info.split('\n'):
                if '지성환' in line or '950924' in line:
                    target_row = int(line.split(':')[0])
                    print(f"\n지성환 행 인덱스: {target_row}")
                    break

        if target_row is None:
            print("\n지성환 못 찾음. 0번 행으로 테스트")
            target_row = 0

        # 납부서 버튼 클릭
        print(f"\n납부서 버튼 클릭 (row {target_row})...")
        clicked = await ev(ws, f"""(function(){{
    var container = document.querySelector('[id*="UTERNAAZ0Z31_wframe"]');
    var rows = Array.from(container.querySelectorAll('table tbody tr'))
        .filter(function(tr){{ return tr.querySelectorAll('td').length >= 13; }});
    var row = rows[{target_row}];
    var tds = Array.from(row.querySelectorAll('td'));
    var cell = tds[13];
    var cellText = cell ? cell.innerText.trim() : 'no_cell';
    if (cellText === '-' || cellText === '') return 'empty:' + cellText;
    var btn = cell ? cell.querySelector('input[type=button],button') : null;
    if (!btn) return 'no_btn:' + cellText;
    btn.click();
    return 'clicked:' + (btn.value||btn.innerText||'?');
}})()""", 3)
        print(f"클릭 결과: {clicked}")

        if not clicked or 'clicked' not in str(clicked):
            print("버튼 없음 — 종료")
            return

        # 새 탭 대기 (최대 15초)
        print("새 탭 대기...")
        new_tab = None
        for _ in range(30):
            await asyncio.sleep(0.5)
            tabs = get_tabs()
            for t in tabs:
                if t['id'] not in known_ids and 'devtools' not in t.get('url',''):
                    new_tab = t
                    break
            if new_tab:
                break

        if new_tab:
            print(f"새 탭 발견: {new_tab['url']}")
            await asyncio.sleep(3)
            # ClipReport 탭 HTML 검사
            async with websockets.connect(new_tab['webSocketDebuggerUrl']) as ws2:
                title = await ev(ws2, "document.title", 1)
                print(f"탭 제목: {title}")
                # 모든 버튼/링크 목록
                btns = await ev(ws2, """(function(){
    var els = Array.from(document.querySelectorAll('input[type=button],button,a[href],.report_menu_button,[class*=pdf],[class*=PDF],[class*=print]'));
    return els.map(function(el){
        return el.tagName + ' cls=[' + el.className.slice(0,60) + '] val=[' + (el.value||el.innerText||'').trim().slice(0,30) + ']';
    }).join('\n') || '(요소 없음)';
})()""", 2)
                print("=== ClipReport 버튼/링크 목록 ===")
                print(btns)
                # body 텍스트 미리보기
                body = await ev(ws2, "document.body ? document.body.innerText.slice(0,200) : 'no body'", 3)
                print(f"\n=== body 미리보기 ===\n{body}")
        else:
            # 팝업이 뜬 경우 — 현재 페이지에서 확인
            print("새 탭 없음 — 팝업이 현재 페이지에 열렸을 가능성")
            popup_info = await ev(ws, """(function(){
    var popups = Array.from(document.querySelectorAll('[class*=popup],[class*=dialog],[class*=modal],div[style*=block]'))
        .filter(function(el){ var r=el.getBoundingClientRect(); return r.width>50&&r.height>50; });
    return popups.map(function(el){
        return el.tagName+'['+el.id+'] cls=['+el.className.slice(0,40)+'] txt=['+el.innerText.trim().slice(0,60)+']';
    }).join('\n') || '(팝업 없음)';
})()""", 4)
            print("=== 현재 페이지 팝업 ===")
            print(popup_info)

asyncio.run(main())
